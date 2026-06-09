/* Lab Check Bot — single-page client.
 *
 * Flow:
 *   1. User pastes a GitHub URL (repo, /tree/<branch>/<path>, /blob/..., or raw).
 *   2. Client fetches markdown (recursively for folders) via api.github.com.
 *   3. Parses steps, detects portals + named entities (agents, pools).
 *   4. User signs in with Microsoft (Graph scopes only).
 *   5. Runs Microsoft Graph checks. Copilot Studio = "Dataverse SP present" via
 *      Graph + manual-verify links per agent. No Dataverse token, ever.
 *   6. Renders results.
 *
 * Configuration: multi-tenant app registration in Entra ID, single-page application
 * platform, redirect URI = current page URL. Default CLIENT_ID below is the dev one
 * — override by appending ?clientId=<guid> to the URL, or hard-coding here.
 */

(() => {
  "use strict";

  // ──────────────────────────────────────────────────────────────────────────
  // Config
  // ──────────────────────────────────────────────────────────────────────────

  const DEFAULT_CLIENT_ID = "34e737f5-9e6c-429c-bd88-c99cf5758253";
  const CLIENT_ID =
    new URLSearchParams(location.search).get("clientId") || DEFAULT_CLIENT_ID;
  const REDIRECT_URI = location.origin + location.pathname;

  const GRAPH_SCOPES = [
    "User.Read",
    "Application.Read.All",
    "Directory.Read.All",
    "RoleManagement.Read.Directory",
    "DeviceManagementServiceConfig.Read.All",
    "DeviceManagementConfiguration.Read.All",
  ];

  // Dataverse discovery scope. Requires Dynamics CRM → user_impersonation
  // delegated permission on the app reg + admin consent. We never auto-redirect
  // for this — silent first, popup on user click only.
  const DATAVERSE_DISCOVERY_SCOPE =
    "https://globaldisco.crm.dynamics.com/.default";

  // Azure Resource Manager scope. Requires Azure Service Management →
  // user_impersonation delegated permission on the app reg + admin consent.
  // Same rule as Dataverse: silent first, popup on user click only.
  const ARM_SCOPE = "https://management.azure.com/user_impersonation";

  // ──────────────────────────────────────────────────────────────────────────
  // State
  // ──────────────────────────────────────────────────────────────────────────

  /** @type {msal.PublicClientApplication|null} */
  let msalInstance = null;
  /** @type {{rawText:string,title:string,url:string,steps:string[],portals:string[],names:{agents:string[],pools:string[]}}|null} */
  let lab = null;

  const LAB_STORAGE_KEY = "labCheckBot.lab.v1";

  function persistLab() {
    try {
      if (lab) localStorage.setItem(LAB_STORAGE_KEY, JSON.stringify(lab));
      else localStorage.removeItem(LAB_STORAGE_KEY);
    } catch { /* quota or private mode — ignore */ }
  }

  function restoreLab() {
    try {
      const raw = localStorage.getItem(LAB_STORAGE_KEY);
      if (!raw) return null;
      const obj = JSON.parse(raw);
      if (obj && obj.rawText && obj.names) {
        // Backfill new name buckets so older cached labs don't crash renderLab.
        obj.names.azure = obj.names.azure || {};
        obj.names.github = obj.names.github || {};
        obj.names.ado = obj.names.ado || {};
        return obj;
      }
    } catch { /* ignore */ }
    return null;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // DOM helpers
  // ──────────────────────────────────────────────────────────────────────────

  const $ = (id) => document.getElementById(id);
  const show = (id) => $(id).classList.remove("hidden");
  const hide = (id) => $(id).classList.add("hidden");

  function showError(msg) {
    const el = $("error-banner");
    if (!msg) {
      el.classList.add("hidden");
      el.textContent = "";
      return;
    }
    el.textContent = msg;
    el.classList.remove("hidden");
  }

  function showInfo(msg) {
    const el = $("info-banner");
    if (!msg) {
      el.classList.add("hidden");
      el.textContent = "";
      return;
    }
    el.textContent = msg;
    el.classList.remove("hidden");
  }

  // ──────────────────────────────────────────────────────────────────────────
  // GitHub lab fetching
  // ──────────────────────────────────────────────────────────────────────────

  const MD_EXTENSIONS = [".md", ".mdx", ".markdown"];

  function isMarkdownPath(p) {
    const lower = p.toLowerCase();
    return MD_EXTENSIONS.some((ext) => lower.endsWith(ext));
  }

  function parseGitHubUrl(input) {
    let url;
    try {
      url = new URL(input);
    } catch {
      return null;
    }
    // raw.githubusercontent.com/<owner>/<repo>/<ref>/<path>
    if (url.hostname === "raw.githubusercontent.com") {
      const parts = url.pathname.split("/").filter(Boolean);
      if (parts.length >= 4) {
        return {
          kind: "blob",
          owner: parts[0],
          repo: parts[1],
          ref: parts[2],
          path: parts.slice(3).join("/"),
        };
      }
    }
    // GitHub Pages: <owner>.github.io/<repo>[/...] → load the backing repo.
    // User/org root site (<owner>.github.io with no path) maps to the
    // special <owner>.github.io repo.
    if (url.hostname.endsWith(".github.io")) {
      const owner = url.hostname.slice(0, -".github.io".length);
      if (!owner) return null;
      const pageParts = url.pathname.split("/").filter(Boolean);
      const repo = pageParts[0] || `${owner}.github.io`;
      return { kind: "tree", owner, repo, ref: "HEAD", path: "" };
    }
    if (url.hostname !== "github.com" && !url.hostname.endsWith(".github.com")) {
      return null;
    }
    const parts = url.pathname.split("/").filter(Boolean);
    if (parts.length < 2) return null;
    const owner = parts[0];
    const repo = parts[1];
    if (parts.length === 2) {
      return { kind: "tree", owner, repo, ref: "HEAD", path: "" };
    }
    const kind = parts[2]; // "tree" or "blob"
    if (kind !== "tree" && kind !== "blob") return null;
    const ref = parts[3] || "HEAD";
    const path = parts.slice(4).join("/");
    return { kind, owner, repo, ref, path };
  }

  async function ghJson(url) {
    const resp = await fetch(url, {
      headers: { Accept: "application/vnd.github+json" },
    });
    if (!resp.ok) {
      let detail = resp.statusText;
      try {
        const j = await resp.json();
        if (j && j.message) detail = j.message;
      } catch { /* ignore */ }
      if (resp.status === 403 && /rate limit/i.test(detail)) {
        throw new Error(
          "GitHub API rate limit hit (60 requests/hour for unauthenticated users). Wait an hour or try a different network."
        );
      }
      throw new Error(`GitHub ${resp.status}: ${detail}`);
    }
    return resp.json();
  }

  async function ghText(url) {
    const resp = await fetch(url);
    if (!resp.ok) {
      throw new Error(`Fetch ${resp.status} for ${url}`);
    }
    return resp.text();
  }

  /**
   * Resolve a ref (branch/tag/HEAD) to a commit SHA so the git tree call works.
   */
  async function resolveRef(owner, repo, ref) {
    if (ref && ref !== "HEAD" && /^[0-9a-f]{40}$/i.test(ref)) return ref;
    // Try default branch when ref is HEAD
    if (!ref || ref === "HEAD") {
      const repoInfo = await ghJson(`https://api.github.com/repos/${owner}/${repo}`);
      ref = repoInfo.default_branch || "main";
    }
    const branch = await ghJson(
      `https://api.github.com/repos/${owner}/${repo}/branches/${encodeURIComponent(ref)}`
    ).catch(() => null);
    if (branch && branch.commit && branch.commit.sha) return branch.commit.sha;
    // Maybe it's a tag
    const tag = await ghJson(
      `https://api.github.com/repos/${owner}/${repo}/git/ref/tags/${encodeURIComponent(ref)}`
    ).catch(() => null);
    if (tag && tag.object && tag.object.sha) return tag.object.sha;
    return ref; // last-ditch: assume caller knows what they're doing
  }

  /**
   * Recursive folder/file fetch via the git/trees API — 1 request for the tree,
   * then N raw fetches for markdown. raw.githubusercontent.com is not rate-limited.
   */
  async function collectMarkdown(owner, repo, ref, path) {
    const sha = await resolveRef(owner, repo, ref);
    const tree = await ghJson(
      `https://api.github.com/repos/${owner}/${repo}/git/trees/${sha}?recursive=1`
    );
    if (tree.truncated) {
      console.warn("Git tree truncated; not all files will be visible.");
    }
    const prefix = path ? path.replace(/\/+$/, "") + "/" : "";
    const want = (tree.tree || []).filter(
      (e) =>
        e.type === "blob" &&
        isMarkdownPath(e.path) &&
        (prefix === "" || e.path === prefix.slice(0, -1) || e.path.startsWith(prefix))
    );
    const out = [];
    for (const entry of want) {
      const raw = `https://raw.githubusercontent.com/${owner}/${repo}/${ref || sha}/${entry.path}`;
      const content = await ghText(raw).catch(() => null);
      if (content !== null) out.push({ path: entry.path, content });
    }
    return out;
  }

  async function fetchLab(input) {
    const parsed = parseGitHubUrl(input);
    if (!parsed) {
      throw new Error("Could not understand that URL. Use a github.com or *.github.io link.");
    }
    if (parsed.kind === "blob") {
      const ref = parsed.ref && parsed.ref !== "HEAD" ? parsed.ref : "HEAD";
      const content = await ghText(
        `https://raw.githubusercontent.com/${parsed.owner}/${parsed.repo}/${ref}/${parsed.path}`
      );
      return {
        title: parsed.path.split("/").pop() || `${parsed.repo}`,
        url: input,
        files: [{ path: parsed.path, content }],
      };
    }
    const files = await collectMarkdown(parsed.owner, parsed.repo, parsed.ref, parsed.path);
    if (files.length === 0) {
      throw new Error(
        "No markdown files found at that location. Double-check the URL points to a folder or file containing .md content."
      );
    }
    files.sort((a, b) => a.path.localeCompare(b.path));
    const title = (parsed.path ? parsed.path.split("/").pop() : parsed.repo) || parsed.repo;
    return { title, url: input, files };
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Lab parsing — steps, portals, named entities
  // ──────────────────────────────────────────────────────────────────────────

  const PORTAL_PATTERNS = [
    { name: "Copilot Studio", patterns: [/copilot\s*studio/i, /copilotstudio\.microsoft\.com/i] },
    { name: "Power Platform Admin", patterns: [/power\s*platform\s*admin/i, /admin\.powerplatform/i, /power\s*automate/i] },
    { name: "Microsoft Entra", patterns: [/microsoft\s*entra/i, /entra\.microsoft\.com/i, /aad|azure\s*ad\b/i] },
    { name: "Microsoft Intune", patterns: [/intune/i] },
    { name: "Azure Portal", patterns: [
        /azure\s*portal/i, /portal\.azure\.com/i, /resource\s*group/i,
        /app\s*service/i, /container\s*app/i, /function\s*app/i,
        /azure\s*sql/i, /cosmos\s*db/i, /storage\s*account/i,
        /key\s*vault/i, /ai\s*foundry/i, /azure\s*openai/i,
        /log\s*analytics/i, /application\s*insights/i, /\baz\s+\w+/i,
      ] },
    { name: "Microsoft 365 Admin", patterns: [/microsoft\s*365\s*admin/i, /admin\.microsoft\.com/i] },
    { name: "SharePoint", patterns: [/sharepoint/i] },
    { name: "Teams Admin", patterns: [/teams\s*admin/i] },
    { name: "GitHub", patterns: [
        /github\s*enterprise/i, /github\s*copilot/i, /github\s*actions/i,
        /\bgithub\.com\b/i, /\bgh\s+(?:auth|repo|workflow|run)\b/i,
        /personal\s*access\s*token/i, /\bpat\b/i,
      ] },
    { name: "Azure DevOps", patterns: [
        /azure\s*devops/i, /\bdev\.azure\.com\b/i, /\bvisualstudio\.com\b/i,
        /\bado\b/i, /azure\s*pipelines?/i, /azure\s*boards?/i, /azure\s*repos?/i,
      ] },
  ];

  function detectPortals(text) {
    const found = new Set();
    for (const { name, patterns } of PORTAL_PATTERNS) {
      if (patterns.some((p) => p.test(text))) found.add(name);
    }
    return Array.from(found);
  }

  // Match quoted, bolded, or unquoted "agent named Foo Bar" forms.
  const AGENT_PATTERNS = [
    /agent\s+named\s+[`*"]+([^`*"\n]{2,80})[`*"]+/gi,
    /create\s+(?:an?\s+)?(?:agent\s+)?(?:called|named)\s+[`*"]+([^`*"\n]{2,80})[`*"]+/gi,
    /(?:^|\s)[`]([A-Z][A-Za-z0-9 _-]{2,60}\s*(?:Agent|Bot))[`]/g,
    /agent\s+named\s+([A-Z][A-Za-z0-9 _-]{2,60}?)(?=\s+(?:in|with|to|that|using|for)\b|[.,])/g,
    /named\s+([A-Z][A-Za-z0-9 _-]{2,60}?\s*(?:Agent|Bot))\b/g,
  ];

  const POOL_PATTERNS = [
    /(?:cloud\s*pc\s+pool|machine\s+pool|machine\s+group)\s+named\s+[`*"]+([^`*"\n]{2,80})[`*"]+/gi,
    /(?:^|\s)[`]([A-Za-z][A-Za-z0-9_-]{2,60}(?:Pool|Group))[`]/g,
    /(?:cloud\s*pc\s+pool|machine\s+pool|machine\s+group)\s+named\s+([A-Za-z][A-Za-z0-9_-]{2,60})\b/g,
  ];

  // Azure resource names — looks for `--name foo`, "named foo", or quoted
  // identifiers next to resource-type keywords.
  const AZURE_RESOURCE_PATTERNS = {
    resourceGroups: [
      /resource\s*group\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9_.-]{2,80})[`*"]+/gi,
      /--resource-group\s+([A-Za-z][A-Za-z0-9_.-]{2,80})/gi,
      /\b-g\s+([A-Za-z][A-Za-z0-9_.-]{2,80})/gi,
    ],
    appServices: [
      /(?:app\s*service|web\s*app)\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,60})[`*"]+/gi,
      /az\s+webapp\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([A-Za-z][A-Za-z0-9-]{2,60})/gi,
    ],
    containerApps: [
      /container\s*app\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,60})[`*"]+/gi,
      /az\s+containerapp\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([A-Za-z][A-Za-z0-9-]{2,60})/gi,
    ],
    functionApps: [
      /function\s*app\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,60})[`*"]+/gi,
      /az\s+functionapp\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([A-Za-z][A-Za-z0-9-]{2,60})/gi,
    ],
    sqlServers: [
      /sql\s*server\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,60})[`*"]+/gi,
      /az\s+sql\s+server\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([A-Za-z][A-Za-z0-9-]{2,60})/gi,
    ],
    storageAccounts: [
      /storage\s*account\s+(?:named\s+)?[`*"]+([a-z0-9]{3,24})[`*"]+/gi,
      /az\s+storage\s+account\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([a-z0-9]{3,24})/gi,
    ],
    keyVaults: [
      /key\s*vault\s+(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,24})[`*"]+/gi,
      /az\s+keyvault\s+\w[\w-]*\s+(?:[^\n]*\s+)?--name\s+([A-Za-z][A-Za-z0-9-]{2,24})/gi,
    ],
    openAiAccounts: [
      /(?:azure\s*open\s*ai|ai\s*foundry|cognitive\s*services)\s+(?:account\s+)?(?:named\s+)?[`*"]+([A-Za-z][A-Za-z0-9-]{2,60})[`*"]+/gi,
    ],
  };

  // GitHub identifiers: org/repo, gh CLI invocations, github.com URLs.
  // The negative lookbehind on the URL pattern excludes `docs.github.com`,
  // `gist.github.com`, etc. so we don't catch documentation paths.
  const GITHUB_PATTERNS = {
    repos: [
      /(?<![A-Za-z0-9.-])github\.com\/([A-Za-z0-9][A-Za-z0-9-]{0,38})\/([A-Za-z0-9._-]{1,100})(?![A-Za-z0-9._-])/g,
      /\bgh\s+repo\s+(?:create|clone|view|fork)\s+([A-Za-z0-9][A-Za-z0-9-]{0,38})\/([A-Za-z0-9._-]{1,100})\b/g,
      /(?:fork|clone)\s+[`*"]+([A-Za-z0-9][A-Za-z0-9-]{0,38})\/([A-Za-z0-9._-]{1,100})[`*"]+/gi,
    ],
    workflows: [
      /\.github\/workflows\/([A-Za-z0-9_.-]+\.ya?ml)/g,
      /workflow\s+(?:named\s+)?[`*"]+([A-Za-z0-9_. -]{3,80})[`*"]+/gi,
    ],
  };

  // GitHub docs URL roots like docs.github.com/en/<area>/... should not be
  // mistaken for real repos. These are the first-segment "areas" we drop.
  const GITHUB_DOCS_OWNERS = new Set([
    "en", "fr", "de", "es", "ja", "ko", "zh", "pt", "ru", "it",
    "account", "actions", "billing", "code-security", "codespaces",
    "copilot", "desktop", "discussions", "enterprise-cloud", "enterprise-server",
    "get-started", "git-guides", "graphql", "issues", "manual", "migrations",
    "organizations", "pages", "pull-requests", "repositories", "rest",
    "search-github", "site-policy", "support", "webhooks",
  ]);

  // Azure DevOps identifiers — org URLs and project names. Require 3+ chars
  // and exclude common docs/path segments so we don't catch e.g. /code, /docs.
  const ADO_PATTERNS = {
    orgs: [
      /\bdev\.azure\.com\/([A-Za-z0-9][A-Za-z0-9-]{2,60})(?=\/|\s|$)/g,
      /\b([A-Za-z0-9][A-Za-z0-9-]{2,60})\.visualstudio\.com\b/g,
    ],
    projects: [
      /\bdev\.azure\.com\/[A-Za-z0-9][A-Za-z0-9-]{2,60}\/([A-Za-z0-9][A-Za-z0-9 _.-]{2,60}?)(?=\/|\s|$)/g,
      /(?:azure\s*devops|ado)\s+project\s+(?:named\s+)?[`*"]+([A-Za-z0-9][A-Za-z0-9 _.-]{2,60})[`*"]+/gi,
    ],
    pipelines: [
      /pipeline\s+(?:named\s+)?[`*"]+([A-Za-z0-9][A-Za-z0-9 _.-]{2,80})[`*"]+/gi,
    ],
  };

  const ADO_STOPWORDS = new Set([
    "code", "docs", "boards", "repos", "pipelines", "wiki",
    "marketplace", "settings", "_git", "_apis", "_build",
  ]);

  // Stopwords that are too generic to be real resource names.
  const RESOURCE_STOPWORDS = new Set([
    "myapp", "mywebapp", "mygroup", "test", "demo", "example", "sample",
    "your-app", "your-rg", "your-name", "<your-app>", "<your-rg>",
    "main", "master", "dev", "prod", "production", "staging",
  ]);

  function extractNamed(text, patternMap) {
    const result = {};
    for (const [key, patterns] of Object.entries(patternMap)) {
      const set = new Set();
      for (const p of patterns) {
        p.lastIndex = 0;
        let m;
        while ((m = p.exec(text)) !== null) {
          // If the regex has two capture groups (owner, repo) glue them.
          let v;
          if (m.length >= 3 && m[2] !== undefined) {
            v = `${(m[1] || "").trim()}/${(m[2] || "").trim()}`.replace(/\.git$/, "");
          } else {
            v = (m[1] || "").trim();
          }
          if (
            v.length >= 2 &&
            v.length <= 100 &&
            !RESOURCE_STOPWORDS.has(v.toLowerCase()) &&
            !v.startsWith("<")
          ) {
            // Repo-specific filter: drop docs.github.com first-segment areas.
            if (key === "repos" && v.includes("/")) {
              const owner = v.split("/")[0].toLowerCase();
              if (GITHUB_DOCS_OWNERS.has(owner)) continue;
            }
            // ADO-specific filter: drop common docs/path segments.
            if ((key === "orgs" || key === "projects") && ADO_STOPWORDS.has(v.toLowerCase())) continue;
            set.add(v);
          }
        }
      }
      if (set.size) result[key] = Array.from(set);
    }
    return result;
  }

  const AGENT_STOPWORDS = new Set([
    "computer-using agent",
    "computer using agent",
    "cua agent",
  ]);

  function extractMatches(text, patterns) {
    const found = new Set();
    for (const p of patterns) {
      p.lastIndex = 0;
      let m;
      while ((m = p.exec(text)) !== null) {
        const v = (m[1] || "").trim();
        if (v.length >= 2 && v.length <= 80) found.add(v);
      }
    }
    return Array.from(found);
  }

  function extractSteps(text) {
    const lines = text.split("\n");
    const steps = [];
    for (const raw of lines) {
      const line = raw.trim();
      if (!line) continue;
      // Headings starting with Task/Exercise/Step
      const head = /^#{1,6}\s+(?:task|exercise|step|part)\s*\d+[:\s.-]+(.+)$/i.exec(line);
      if (head) {
        steps.push(cleanMd(head[1]));
        continue;
      }
      // Numbered list
      const num = /^\d+[.)]\s+(.+)$/.exec(line);
      if (num) {
        const t = cleanMd(num[1]);
        if (t.length > 8 && t.length < 320) steps.push(t);
        continue;
      }
      // Bullet under Success criteria etc.
      const bul = /^[-*+]\s+(.+)$/.exec(line);
      if (bul) {
        const t = cleanMd(bul[1]);
        if (t.length > 8 && t.length < 320) steps.push(t);
      }
    }
    return Array.from(new Set(steps));
  }

  function cleanMd(s) {
    return s
      .replace(/!\[[^\]]*\]\([^)]*\)/g, "")
      .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
      .replace(/[*_`]+/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function analyzeLab(fetched) {
    const combined = fetched.files
      .map((f) => `## Source: ${f.path}\n\n${f.content}`)
      .join("\n\n");

    const steps = extractSteps(combined);
    const portals = detectPortals(combined);
    const agentsRaw = extractMatches(combined, AGENT_PATTERNS);
    const poolsRaw = extractMatches(combined, POOL_PATTERNS);

    const agents = agentsRaw.filter((a) => !AGENT_STOPWORDS.has(a.toLowerCase()));
    const pools = poolsRaw;

    const azure = extractNamed(combined, AZURE_RESOURCE_PATTERNS);
    const github = extractNamed(combined, GITHUB_PATTERNS);
    const ado = extractNamed(combined, ADO_PATTERNS);

    return {
      title: fetched.title,
      url: fetched.url,
      rawText: combined,
      steps,
      portals,
      names: { agents, pools, azure, github, ado },
    };
  }

  // ──────────────────────────────────────────────────────────────────────────
  // MSAL
  // ──────────────────────────────────────────────────────────────────────────

  async function buildMsal() {
    if (typeof msal === "undefined") {
      throw new Error("MSAL library failed to load. Try disabling ad blockers or InPrivate mode.");
    }
    const inst = new msal.PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/organizations",
        redirectUri: REDIRECT_URI,
        navigateToLoginRequestUrl: true,
      },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true },
    });
    if (typeof inst.initialize === "function") {
      await inst.initialize();
    }
    // Pick up sign-in / consent redirects.
    await inst.handleRedirectPromise().catch(() => null);
    return inst;
  }

  async function getMsal() {
    if (!msalInstance) msalInstance = await buildMsal();
    return msalInstance;
  }

  function getActiveAccount() {
    if (!msalInstance) return null;
    return msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
  }

  async function signIn() {
    const inst = await getMsal();
    await inst.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  }

  async function switchTenant() {
    const inst = await getMsal();
    inst.setActiveAccount(null);
    await inst.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  }

  function signOut() {
    try { localStorage.removeItem(LAB_STORAGE_KEY); } catch { /* ignore */ }
    if (!msalInstance) {
      localStorage.clear();
      sessionStorage.clear();
      location.reload();
      return;
    }
    msalInstance.logoutRedirect({ postLogoutRedirectUri: REDIRECT_URI });
  }

  async function getToken(scopes, { forceRefresh = false } = {}) {
    const inst = await getMsal();
    const account = getActiveAccount();
    if (!account) throw new Error("Not signed in.");
    const resp = await inst.acquireTokenSilent({ scopes, account, forceRefresh });
    return resp.accessToken;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Tenant checks
  // ──────────────────────────────────────────────────────────────────────────

  async function graphGet(path, beta = false, { forceRefresh = false } = {}) {
    const token = await getToken(["https://graph.microsoft.com/.default"], { forceRefresh }).catch(
      async () => getToken(GRAPH_SCOPES, { forceRefresh })
    );
    const base = beta
      ? "https://graph.microsoft.com/beta"
      : "https://graph.microsoft.com/v1.0";
    const resp = await fetch(base + path, {
      headers: { Authorization: "Bearer " + token },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      throw new Error(`Graph ${resp.status}: ${body || resp.statusText}`);
    }
    return resp.json();
  }

  const STATIC_CHECKS = [
    {
      id: "w365-sp",
      name: "Windows 365 service principal",
      portals: ["Microsoft Entra", "Copilot Studio"],
      keywords: [/windows\s*365/i, /service\s*principal/i],
      run: async () => {
        const d = await graphGet(
          "/servicePrincipals?$filter=appId eq '0af06dc6-e4b5-4f28-818e-e78e62d137a5'&$select=id,displayName"
        );
        return (d.value || []).length
          ? { status: "pass", msg: `Found: ${d.value[0].displayName}` }
          : { status: "fail", msg: "Not found." };
      },
    },
    {
      id: "avd-arm-sp",
      name: "Azure Virtual Desktop ARM provider SP",
      portals: ["Azure Portal", "Microsoft Entra"],
      keywords: [/azure\s*virtual\s*desktop/i, /\bavd\b/i, /service\s*principal/i],
      run: async () => {
        const d = await graphGet(
          "/servicePrincipals?$filter=appId eq '50e95039-b200-4007-bc97-8d5790743a63'&$select=id,displayName"
        );
        return (d.value || []).length
          ? { status: "pass", msg: `Found: ${d.value[0].displayName}` }
          : { status: "fail", msg: "Not found." };
      },
    },
    {
      id: "rd-sp",
      name: "Microsoft Remote Desktop SP",
      portals: ["Azure Portal", "Microsoft Entra"],
      keywords: [/remote\s*desktop/i, /\brdp\b/i],
      run: async () => {
        const d = await graphGet(
          "/servicePrincipals?$filter=appId eq 'a4a365df-50f1-4397-bc59-1a1564b8bb9c'&$select=id,displayName"
        );
        return (d.value || []).length
          ? { status: "pass", msg: `Found: ${d.value[0].displayName}` }
          : { status: "fail", msg: "Not found." };
      },
    },
    {
      id: "entra-p1",
      name: "Entra ID P1/P2 license",
      portals: ["Microsoft Entra"],
      keywords: [/entra/i, /license/i],
      run: async () => {
        const d = await graphGet("/subscribedSkus?$select=skuPartNumber,servicePlans");
        const rows = d.value || [];
        const has = rows.some((r) =>
          (r.servicePlans || []).some(
            (p) =>
              (p.servicePlanName === "AAD_PREMIUM" || p.servicePlanName === "AAD_PREMIUM_P2") &&
              (p.provisioningStatus === "Success" || p.provisioningStatus === "PendingInput")
          )
        );
        return has
          ? { status: "pass", msg: "P1/P2 service plan detected." }
          : { status: "warn", msg: "P1/P2 not confirmed." };
      },
    },
    {
      id: "me-admin",
      name: "Signed-in user has admin role",
      portals: [],
      keywords: [/admin/i],
      run: async () => {
        const d = await graphGet("/me/memberOf?$select=displayName,roleTemplateId");
        const adminIds = new Set([
          "62e90394-69f5-4237-9190-012177145e10", // Global admin
          "158c047a-c907-4556-b7ef-446551a6b5f7", // Cloud app admin
          "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3", // App admin
          "3a2c62db-5318-420d-8d74-23affee5d9d5", // Intune admin
        ]);
        const hits = (d.value || []).filter((r) => adminIds.has(r.roleTemplateId));
        return hits.length
          ? { status: "pass", msg: `Roles: ${hits.map((r) => r.displayName).join(", ")}` }
          : { status: "warn", msg: "No common admin role found for signed-in user." };
      },
    },
    {
      id: "intune-enroll",
      name: "Intune enrollment configurations exist",
      portals: ["Microsoft Intune"],
      keywords: [/intune/i, /enrollment/i],
      run: async () => {
        const d = await graphGet("/deviceManagement/deviceEnrollmentConfigurations?$select=id,displayName", true);
        const n = (d.value || []).length;
        return n
          ? { status: "pass", msg: `${n} configuration(s) found.` }
          : { status: "warn", msg: "Could not read enrollment configurations." };
      },
    },
  ];

  function checkApplies(check, lab) {
    if (check.portals && check.portals.length > 0) {
      const overlap = check.portals.some((p) => lab.portals.includes(p));
      if (!overlap) return false;
    }
    if (check.keywords && check.keywords.length > 0) {
      const txt = lab.rawText;
      const hits = check.keywords.filter((k) => k.test(txt));
      // require at least one keyword match if keywords specified
      if (hits.length === 0) return false;
    }
    return true;
  }

  // ── Copilot Studio / Power Platform (Graph-only — no Dataverse token needed)
  //
  // The Dataverse service principal (appId 00000007-0000-0000-c000-000000000000)
  // being present in the tenant indicates Power Platform / Copilot Studio is
  // provisioned. We deliberately avoid acquiring a Dataverse token because that
  // would require Dynamics CRM user_impersonation delegated permission on the
  // app reg + admin consent in every tenant we're used in — defeating the
  // "works in any CDX tenant out of the box" goal.

  async function checkDataverseSp() {
    try {
      const d = await graphGet(
        "/servicePrincipals?$filter=appId eq '00000007-0000-0000-c000-000000000000'&$select=id,displayName,accountEnabled"
      );
      const sp = (d.value || [])[0];
      if (!sp) {
        return {
          status: "fail",
          msg: "Dataverse service principal not found. Create a Power Platform environment with Dataverse before running Copilot Studio steps.",
        };
      }
      if (sp.accountEnabled === false) {
        return { status: "fail", msg: `Dataverse SP "${sp.displayName}" is disabled.` };
      }
      return { status: "pass", msg: `Dataverse SP active: ${sp.displayName}` };
    } catch (e) {
      return { status: "warn", msg: `Could not check Dataverse SP: ${e.message}` };
    }
  }

  async function runCopilotStudioChecks(lab) {
    const needsCs =
      (lab.portals || []).includes("Copilot Studio") ||
      (lab.portals || []).includes("Power Platform Admin") ||
      (lab.names.agents || []).length > 0;
    if (!needsCs) return [];

    const results = [];
    results.push({
      id: "cs-dataverse-sp",
      name: "Copilot Studio / Dataverse provisioned",
      ...(await checkDataverseSp()),
    });

    const agentNames = lab.names.agents || [];
    if (agentNames.length === 0) return results;

    // Try silent Dataverse token. If it fails (no consent / no permission),
    // emit one SKIP result with a "Grant access" button — never auto-redirect.
    let instances = null;
    let dvError = null;
    try {
      instances = await discoverInstances();
    } catch (e) {
      dvError = e;
    }

    if (!instances) {
      results.push({
        id: "cs-grant-dataverse",
        name: "Verify agents in Copilot Studio (Dataverse access required)",
        status: "skip",
        msg: `Click below to grant Dataverse access (one-time per tenant), then re-run checks.`,
        action: {
          label: "🔓 Grant Dataverse access",
          handler: "grantDataverse",
        },
      });
      for (const name of agentNames) {
        results.push({
          id: `cs-agent-${slug(name)}`,
          name: `Agent "${name}"`,
          status: "skip",
          msg: "Pending Dataverse access — see action above.",
        });
      }
      return results;
    }

    // Gather all bots across all envs.
    const allBots = [];
    for (const inst of instances) {
      try {
        const bots = await listBots(inst);
        allBots.push(...bots);
      } catch {
        /* ignore env */
      }
    }

    for (const name of agentNames) {
      const targetTokens = significantTokens(name);
      const matches = allBots.filter((b) => {
        const candidate = significantTokens(b.name || "");
        return targetTokens.some((t) => candidate.includes(t));
      });
      const active = matches.filter((b) => (b.statecode ?? 0) === 0);
      if (active.length) {
        results.push({
          id: `cs-agent-${slug(name)}`,
          name: `Agent "${name}"`,
          status: "pass",
          msg: `Matched: ${active.map((b) => `"${b.name}" in ${b.env}`).join("; ")}`,
        });
      } else if (matches.length) {
        results.push({
          id: `cs-agent-${slug(name)}`,
          name: `Agent "${name}"`,
          status: "warn",
          msg: `Found but inactive: ${matches.map((b) => `"${b.name}" in ${b.env}`).join("; ")}`,
        });
      } else {
        results.push({
          id: `cs-agent-${slug(name)}`,
          name: `Agent "${name}"`,
          status: "fail",
          msg: "Not found in any accessible Power Platform environment.",
        });
      }
    }
    return results;
  }

  async function discoverInstances() {
    const inst = await getMsal();
    const account = getActiveAccount();
    if (!account) throw new Error("Not signed in");
    // Silent only — caller decides what to do on failure.
    const tokenResp = await inst.acquireTokenSilent({
      scopes: [DATAVERSE_DISCOVERY_SCOPE],
      account,
    });
    const resp = await fetch(
      "https://globaldisco.crm.dynamics.com/api/discovery/v2.0/Instances",
      { headers: { Authorization: "Bearer " + tokenResp.accessToken, Accept: "application/json" } }
    );
    if (!resp.ok) throw new Error(`Discovery ${resp.status}: ${resp.statusText}`);
    const data = await resp.json();
    return data.value || [];
  }

  async function listBots(instance) {
    const orgUrl = instance.ApiUrl.replace(/\/$/, "");
    const inst = await getMsal();
    const account = getActiveAccount();
    const tokenResp = await inst.acquireTokenSilent({
      scopes: [`${orgUrl}/.default`],
      account,
    });
    const resp = await fetch(
      `${orgUrl}/api/data/v9.2/bots?$select=name,statecode&$top=200`,
      { headers: { Authorization: "Bearer " + tokenResp.accessToken, Accept: "application/json" } }
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    return (data.value || []).map((b) => ({
      name: b.name,
      statecode: b.statecode,
      env: instance.FriendlyName || orgUrl,
    }));
  }

  const PP_STOPWORDS = new Set([
    "agent", "agents", "bot", "bots", "copilot", "studio",
    "the", "and", "for", "with", "test",
  ]);
  function significantTokens(s) {
    return s
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter((t) => t.length >= 4 && !PP_STOPWORDS.has(t));
  }

  // Exposed on window so the in-result "Grant" button can call it.
  async function grantDataverse() {
    showError("");
    const account = getActiveAccount();
    const tenantId =
      (account && (account.tenantId || (account.idTokenClaims && account.idTokenClaims.tid))) ||
      "organizations";
    // Open Microsoft's standard admin-consent page in a new tab. Global Admins
    // see a normal Microsoft consent screen (not a blank popup). After they
    // click Accept, AAD redirects to our redirectUri — we just show a small
    // "all done" page there. The user then comes back to this tab and re-runs.
    const redirectUri =
      location.origin +
      location.pathname.replace(/[^/]*$/, "") +
      "auth-redirect.html";
    const url =
      `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/v2.0/adminconsent` +
      `?client_id=${encodeURIComponent(CLIENT_ID)}` +
      `&scope=${encodeURIComponent("https://globaldisco.crm.dynamics.com/user_impersonation")}` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&state=lab-check-bot`;
    showInfo(
      "Opening Microsoft admin-consent page in a new tab. After you click Accept, come back here and click Run checks again."
    );
    window.open(url, "_blank", "noopener");
  }
  window.__labCheckBot = Object.assign(window.__labCheckBot || {}, { grantDataverse });

  // ──────────────────────────────────────────────────────────────────────────
  // Azure Resource Manager (ARM) checks
  // ──────────────────────────────────────────────────────────────────────────

  // Maps lab.names.azure keys → ARM resource-type strings.
  // resourceGroups is special (different endpoint).
  const ARM_RESOURCE_TYPES = {
    appServices: "Microsoft.Web/sites",
    functionApps: "Microsoft.Web/sites", // function apps are Web/sites with kind=functionapp
    containerApps: "Microsoft.App/containerApps",
    sqlServers: "Microsoft.Sql/servers",
    storageAccounts: "Microsoft.Storage/storageAccounts",
    keyVaults: "Microsoft.KeyVault/vaults",
    openAiAccounts: "Microsoft.CognitiveServices/accounts",
  };

  async function armGet(path) {
    const inst = await getMsal();
    const account = getActiveAccount();
    if (!account) throw new Error("Not signed in");
    const tokenResp = await inst.acquireTokenSilent({
      scopes: [ARM_SCOPE],
      account,
    });
    const url = path.startsWith("http")
      ? path
      : `https://management.azure.com${path}`;
    const resp = await fetch(url, {
      headers: {
        Authorization: "Bearer " + tokenResp.accessToken,
        Accept: "application/json",
      },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      throw new Error(`ARM ${resp.status}: ${resp.statusText}${body ? " — " + body.slice(0, 200) : ""}`);
    }
    return resp.json();
  }

  async function listAllSubscriptions() {
    const data = await armGet("/subscriptions?api-version=2020-01-01");
    return (data.value || [])
      .filter((s) => (s.state || "").toLowerCase() === "enabled")
      .map((s) => ({ id: s.subscriptionId, name: s.displayName }));
  }

  async function listResourcesOfType(subId, typeString) {
    // Single-page list (default 100 items). Good enough for lab tenants.
    const url =
      `/subscriptions/${encodeURIComponent(subId)}/resources` +
      `?api-version=2021-04-01&$filter=${encodeURIComponent(`resourceType eq '${typeString}'`)}` +
      `&$top=200`;
    const data = await armGet(url);
    return data.value || [];
  }

  async function listResourceGroups(subId) {
    const url = `/subscriptions/${encodeURIComponent(subId)}/resourcegroups?api-version=2021-04-01&$top=200`;
    const data = await armGet(url);
    return data.value || [];
  }

  async function runAzureChecks(lab) {
    const az = (lab.names && lab.names.azure) || {};
    const needsAzure =
      (lab.portals || []).includes("Azure Portal") ||
      Object.values(az).some((arr) => Array.isArray(arr) && arr.length > 0);
    if (!needsAzure) return [];

    const results = [];

    // Silent ARM token probe. If consent / permission isn't there, emit one
    // SKIP with a "Grant Azure access" button — never auto-redirect.
    let subs = null;
    let armError = null;
    try {
      subs = await listAllSubscriptions();
    } catch (e) {
      armError = e;
    }

    if (!subs) {
      results.push({
        id: "az-grant-arm",
        name: "Verify Azure resources (Azure Resource Manager access required)",
        status: "skip",
        msg:
          "Click below to grant Azure Resource Manager access (one-time per tenant), then re-run checks. " +
          "Also requires the 'Azure Service Management → user_impersonation' delegated permission on the app registration." +
          (armError ? ` (${armError.message})` : ""),
        action: {
          label: "🔓 Grant Azure access",
          handler: "grantAzureArm",
        },
      });
      // Still mark each detected name as pending so the user sees what we
      // would check.
      for (const [key, names] of Object.entries(az)) {
        for (const name of names || []) {
          results.push({
            id: `az-${key}-${slug(name)}`,
            name: `Azure ${humanizeArmKey(key)} "${name}"`,
            status: "skip",
            msg: "Pending Azure access — see action above.",
          });
        }
      }
      return results;
    }

    if (subs.length === 0) {
      results.push({
        id: "az-no-subs",
        name: "Azure subscription available",
        status: "fail",
        msg: "No enabled Azure subscriptions found on this account. The lab likely requires you to create or be assigned a subscription before running its Azure steps.",
      });
      return results;
    }

    results.push({
      id: "az-subs",
      name: "Azure subscriptions accessible",
      status: "pass",
      msg: `Found ${subs.length} enabled subscription(s): ${subs.map((s) => s.name).join(", ")}`,
    });

    // Resource groups (single special-case endpoint)
    if ((az.resourceGroups || []).length > 0) {
      const allRgs = [];
      for (const sub of subs) {
        try {
          const rgs = await listResourceGroups(sub.id);
          allRgs.push(...rgs.map((r) => ({ name: r.name, sub: sub.name })));
        } catch {
          /* ignore one sub failing */
        }
      }
      for (const target of az.resourceGroups) {
        const hit = allRgs.find((r) => r.name.toLowerCase() === target.toLowerCase());
        results.push({
          id: `az-rg-${slug(target)}`,
          name: `Azure resource group "${target}"`,
          status: hit ? "pass" : "fail",
          msg: hit
            ? `Found in subscription "${hit.sub}".`
            : "Not found in any accessible subscription. Create the resource group as the lab describes.",
        });
      }
    }

    // Typed resources
    for (const [key, typeString] of Object.entries(ARM_RESOURCE_TYPES)) {
      const targets = az[key] || [];
      if (targets.length === 0) continue;

      const found = [];
      for (const sub of subs) {
        try {
          const items = await listResourcesOfType(sub.id, typeString);
          for (const r of items) {
            // Function apps are Web/sites with kind containing 'functionapp'.
            if (key === "functionApps") {
              if (!(r.kind || "").toLowerCase().includes("functionapp")) continue;
            } else if (key === "appServices" && typeString === "Microsoft.Web/sites") {
              // Exclude function apps from the regular App Service bucket.
              if ((r.kind || "").toLowerCase().includes("functionapp")) continue;
            }
            found.push({ name: r.name, sub: sub.name });
          }
        } catch {
          /* ignore one sub failing */
        }
      }

      for (const target of targets) {
        const hit = found.find((f) => f.name.toLowerCase() === target.toLowerCase());
        results.push({
          id: `az-${key}-${slug(target)}`,
          name: `Azure ${humanizeArmKey(key)} "${target}"`,
          status: hit ? "pass" : "fail",
          msg: hit
            ? `Found in subscription "${hit.sub}".`
            : `Not found in any accessible subscription. The lab expected an existing ${humanizeArmKey(key)} with this name.`,
        });
      }
    }

    return results;
  }

  function humanizeArmKey(key) {
    const map = {
      resourceGroups: "resource group",
      appServices: "App Service",
      functionApps: "Function App",
      containerApps: "Container App",
      sqlServers: "SQL server",
      storageAccounts: "storage account",
      keyVaults: "key vault",
      openAiAccounts: "Azure OpenAI / Cognitive Services account",
    };
    return map[key] || key;
  }

  async function grantAzureArm() {
    showError("");
    const account = getActiveAccount();
    const tenantId =
      (account && (account.tenantId || (account.idTokenClaims && account.idTokenClaims.tid))) ||
      "organizations";
    const redirectUri =
      location.origin +
      location.pathname.replace(/[^/]*$/, "") +
      "auth-redirect.html";
    const url =
      `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/v2.0/adminconsent` +
      `?client_id=${encodeURIComponent(CLIENT_ID)}` +
      `&scope=${encodeURIComponent("https://management.azure.com/user_impersonation")}` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&state=lab-check-bot-arm`;
    showInfo(
      "Opening Microsoft admin-consent page for Azure Resource Manager. " +
      "If the app registration doesn't yet have the 'Azure Service Management → user_impersonation' " +
      "delegated permission, add it in Entra → App registrations first, then click this button again."
    );
    window.open(url, "_blank", "noopener");
  }
  window.__labCheckBot = Object.assign(window.__labCheckBot || {}, { grantAzureArm });

  // ──────────────────────────────────────────────────────────────────────────
  // GitHub PAT checks
  // ──────────────────────────────────────────────────────────────────────────

  const GH_PAT_KEY = "labCheckBot.githubPat";
  const ADO_PAT_KEY = "labCheckBot.adoPat";
  const ADO_ORG_KEY = "labCheckBot.adoOrg";

  function getGitHubPat() {
    try { return localStorage.getItem(GH_PAT_KEY) || ""; } catch { return ""; }
  }
  function setGitHubPat(v) {
    try { v ? localStorage.setItem(GH_PAT_KEY, v) : localStorage.removeItem(GH_PAT_KEY); } catch { /* ignore */ }
  }
  function getAdoPat() {
    try { return localStorage.getItem(ADO_PAT_KEY) || ""; } catch { return ""; }
  }
  function setAdoPat(v) {
    try { v ? localStorage.setItem(ADO_PAT_KEY, v) : localStorage.removeItem(ADO_PAT_KEY); } catch { /* ignore */ }
  }
  function getAdoOrg() {
    try { return localStorage.getItem(ADO_ORG_KEY) || ""; } catch { return ""; }
  }
  function setAdoOrg(v) {
    try { v ? localStorage.setItem(ADO_ORG_KEY, v) : localStorage.removeItem(ADO_ORG_KEY); } catch { /* ignore */ }
  }

  async function ghGet(path, pat) {
    const url = path.startsWith("http") ? path : `https://api.github.com${path}`;
    const resp = await fetch(url, {
      headers: {
        Authorization: "Bearer " + pat,
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
      },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      const err = new Error(`GitHub ${resp.status}: ${resp.statusText}${body ? " — " + body.slice(0, 200) : ""}`);
      err.status = resp.status;
      throw err;
    }
    return resp.json();
  }

  async function runGitHubChecks(lab) {
    const gh = (lab.names && lab.names.github) || {};
    const repos = gh.repos || [];
    const workflows = gh.workflows || [];
    const needsGh =
      (lab.portals || []).includes("GitHub") || repos.length > 0 || workflows.length > 0;
    if (!needsGh) return [];

    const results = [];
    const pat = getGitHubPat();

    if (!pat) {
      results.push({
        id: "gh-no-pat",
        name: "GitHub repo / workflow verification",
        status: "skip",
        msg: "No GitHub PAT saved. Expand 'Optional: GitHub & Azure DevOps access tokens' above and paste a PAT (scopes: repo, read:org, workflow) to enable repo and workflow checks.",
      });
      for (const r of repos) {
        results.push({
          id: `gh-repo-${slug(r)}`,
          name: `GitHub repo "${r}"`,
          status: "skip",
          msg: "Pending GitHub PAT — see action above.",
        });
      }
      return results;
    }

    // Validate PAT
    try {
      const me = await ghGet("/user", pat);
      results.push({
        id: "gh-pat",
        name: "GitHub PAT valid",
        status: "pass",
        msg: `Authenticated as @${me.login}.`,
      });
    } catch (e) {
      results.push({
        id: "gh-pat",
        name: "GitHub PAT valid",
        status: "fail",
        msg: `PAT failed: ${e.message}. Re-paste a fresh token.`,
      });
      return results;
    }

    // Per-repo existence + workflow lookup
    for (const r of repos) {
      let repoData = null;
      try {
        repoData = await ghGet(`/repos/${r}`, pat);
        results.push({
          id: `gh-repo-${slug(r)}`,
          name: `GitHub repo "${r}"`,
          status: "pass",
          msg: `Found · ${repoData.private ? "private" : "public"} · default branch ${repoData.default_branch}.`,
        });
      } catch (e) {
        results.push({
          id: `gh-repo-${slug(r)}`,
          name: `GitHub repo "${r}"`,
          status: e.status === 404 ? "fail" : "warn",
          msg: e.status === 404
            ? "Repo not found (or PAT lacks access). The lab may expect you to fork or clone this repo."
            : e.message,
        });
        continue;
      }

      // Check workflows on this repo
      if (workflows.length > 0) {
        let wfData = null;
        try {
          wfData = await ghGet(`/repos/${r}/actions/workflows`, pat);
        } catch {
          /* ignore */
        }
        const wfList = (wfData && wfData.workflows) || [];
        for (const wf of workflows) {
          const hit = wfList.find(
            (w) =>
              (w.path || "").toLowerCase().endsWith(wf.toLowerCase()) ||
              (w.name || "").toLowerCase() === wf.toLowerCase()
          );
          results.push({
            id: `gh-wf-${slug(r)}-${slug(wf)}`,
            name: `GitHub workflow "${wf}" in ${r}`,
            status: hit ? "pass" : "fail",
            msg: hit
              ? `Found · last update ${hit.updated_at || "?"} · state ${hit.state || "?"}.`
              : "Workflow not found in this repo. The lab expected this workflow to exist.",
          });
        }
      }
    }

    return results;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Azure DevOps PAT checks
  // ──────────────────────────────────────────────────────────────────────────

  async function adoGet(org, path, pat) {
    const url = path.startsWith("http")
      ? path
      : `https://dev.azure.com/${encodeURIComponent(org)}${path}`;
    // ADO basic auth: ':pat' base64
    const auth = btoa(":" + pat);
    const resp = await fetch(url, {
      headers: { Authorization: "Basic " + auth, Accept: "application/json" },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      const err = new Error(`ADO ${resp.status}: ${resp.statusText}${body ? " — " + body.slice(0, 200) : ""}`);
      err.status = resp.status;
      throw err;
    }
    return resp.json();
  }

  async function runAdoChecks(lab) {
    const ado = (lab.names && lab.names.ado) || {};
    const orgs = ado.orgs || [];
    const projects = ado.projects || [];
    const pipelines = ado.pipelines || [];
    const needsAdo =
      (lab.portals || []).includes("Azure DevOps") ||
      orgs.length + projects.length + pipelines.length > 0;
    if (!needsAdo) return [];

    const results = [];
    const pat = getAdoPat();
    const orgOverride = getAdoOrg();
    const targetOrg = orgOverride || orgs[0] || null;

    if (!pat || !targetOrg) {
      results.push({
        id: "ado-no-pat",
        name: "Azure DevOps verification",
        status: "skip",
        msg:
          !pat
            ? "No Azure DevOps PAT saved. Expand 'Optional: GitHub & Azure DevOps access tokens' above and paste a PAT (scopes: Project & Team Read, Code Read, Build Read)."
            : "No ADO organization specified. Add the org name in the PAT panel above (or detect one in the lab text).",
      });
      return results;
    }

    // Validate connection by listing projects
    let projectList = [];
    try {
      const data = await adoGet(targetOrg, "/_apis/projects?api-version=7.1&$top=200", pat);
      projectList = data.value || [];
      results.push({
        id: "ado-org",
        name: `ADO organization "${targetOrg}"`,
        status: "pass",
        msg: `Connected · ${projectList.length} project(s) visible.`,
      });
    } catch (e) {
      results.push({
        id: "ado-org",
        name: `ADO organization "${targetOrg}"`,
        status: "fail",
        msg: `Could not list projects: ${e.message}. Check PAT scopes and org name.`,
      });
      return results;
    }

    // Per-project existence
    for (const p of projects) {
      const hit = projectList.find(
        (x) => (x.name || "").toLowerCase() === p.toLowerCase()
      );
      results.push({
        id: `ado-project-${slug(p)}`,
        name: `ADO project "${p}"`,
        status: hit ? "pass" : "fail",
        msg: hit
          ? `Found in "${targetOrg}" · state ${hit.state || "?"}.`
          : `Not found in org "${targetOrg}". The lab expected a project with this name.`,
      });
    }

    // Per-pipeline existence (search across known projects)
    for (const pipeName of pipelines) {
      let found = null;
      for (const proj of projectList) {
        try {
          const data = await adoGet(
            targetOrg,
            `/${encodeURIComponent(proj.name)}/_apis/build/definitions?api-version=7.1&name=${encodeURIComponent(pipeName)}`,
            pat
          );
          if ((data.value || []).length > 0) {
            found = { project: proj.name, def: data.value[0] };
            break;
          }
        } catch {
          /* ignore */
        }
      }
      results.push({
        id: `ado-pipeline-${slug(pipeName)}`,
        name: `ADO pipeline "${pipeName}"`,
        status: found ? "pass" : "fail",
        msg: found
          ? `Found in project "${found.project}".`
          : `Not found in any visible project under "${targetOrg}".`,
      });
    }

    return results;
  }

  function slug(s) {
    return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
  }

  // For a failing/warning check, surface the lab step(s) the user was supposed
  // to perform. Pure regex / token-overlap matching against parsed lab.steps —
  // no LLM, no Graph follow-up. Returns up to `limit` step strings (snippets).
  function findRelevantSteps(result, labObj, limit = 3) {
    if (!labObj || !Array.isArray(labObj.steps) || labObj.steps.length === 0) return [];

    // Search terms come from:
    //  1. Quoted strings in the check name/msg ("Invoice Processing Agent")
    //  2. Significant tokens from the check name (drops portal stopwords)
    //  3. Detected named items (agents, pools) if the check id mentions them
    const haystack = `${result.name || ""} ${result.msg || ""}`;
    const phrases = [];
    const phraseRe = /"([^"]+)"|“([^”]+)”/g;
    let m;
    while ((m = phraseRe.exec(haystack))) {
      const p = (m[1] || m[2] || "").trim();
      if (p.length >= 3) phrases.push(p);
    }
    const tokens = new Set(significantTokens(result.name || ""));
    // Names referenced by id (e.g. cs-agent-<slug>)
    const named = [
      ...(labObj.names?.agents || []),
      ...(labObj.names?.pools || []),
    ];
    for (const n of named) {
      if (result.id && result.id.includes(slug(n))) phrases.push(n);
    }
    if (phrases.length === 0 && tokens.size === 0) return [];

    const phrasesLower = phrases.map((p) => p.toLowerCase());
    const scored = [];
    labObj.steps.forEach((step, idx) => {
      const lower = step.toLowerCase();
      let score = 0;
      for (const p of phrasesLower) if (lower.includes(p)) score += 5;
      for (const t of tokens) if (lower.includes(t)) score += 1;
      if (score > 0) scored.push({ idx, step, score });
    });
    scored.sort((a, b) => b.score - a.score || a.idx - b.idx);
    return scored.slice(0, limit).map((s) => s.step);
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Rendering
  // ──────────────────────────────────────────────────────────────────────────

  function renderLab() {
    if (!lab) return;
    $("lab-title").textContent = lab.title;
    $("lab-meta").innerHTML = `<a href="${lab.url}" target="_blank" rel="noopener">${lab.url}</a> · ${lab.steps.length} step(s) parsed`;

    const portals = $("portals");
    portals.innerHTML = "";
    if (lab.portals.length === 0) {
      portals.innerHTML = `<span class="muted">No portals detected.</span>`;
    } else {
      lab.portals.forEach((p) => {
        const tag = document.createElement("span");
        tag.className = "portal-tag";
        tag.textContent = p;
        portals.appendChild(tag);
      });
    }

    const named = $("named-items");
    const az = lab.names.azure || {};
    const gh = lab.names.github || {};
    const ado = lab.names.ado || {};
    const hasAnyName =
      lab.names.agents.length ||
      lab.names.pools.length ||
      Object.values(az).some((v) => v && v.length) ||
      Object.values(gh).some((v) => v && v.length) ||
      Object.values(ado).some((v) => v && v.length);
    if (!hasAnyName) {
      named.innerHTML = `<p class="muted">No named resources detected.</p>`;
    } else {
      const row = (label, items) =>
        items && items.length
          ? `<div style="margin-top:6px;"><strong>${escape(label)}:</strong> ${items.map((a) => `<code>${escape(a)}</code>`).join(", ")}</div>`
          : "";
      const parts = [];
      if (lab.names.agents.length) parts.push(row("Agents", lab.names.agents));
      if (lab.names.pools.length) parts.push(row("Pools/Groups", lab.names.pools));
      // Azure
      parts.push(row("Azure resource groups", az.resourceGroups));
      parts.push(row("Azure App Services / Web Apps", az.appServices));
      parts.push(row("Azure Container Apps", az.containerApps));
      parts.push(row("Azure Function Apps", az.functionApps));
      parts.push(row("Azure SQL servers", az.sqlServers));
      parts.push(row("Azure Storage accounts", az.storageAccounts));
      parts.push(row("Azure Key Vaults", az.keyVaults));
      parts.push(row("Azure OpenAI / AI Foundry", az.openAiAccounts));
      // GitHub
      parts.push(row("GitHub repos", gh.repos));
      parts.push(row("GitHub workflows", gh.workflows));
      // Azure DevOps
      parts.push(row("Azure DevOps orgs", ado.orgs));
      parts.push(row("Azure DevOps projects", ado.projects));
      parts.push(row("Azure DevOps pipelines", ado.pipelines));
      named.innerHTML = parts.filter(Boolean).join("");
    }

    const list = $("step-list");
    list.innerHTML = "";
    lab.steps.slice(0, 100).forEach((s) => {
      const li = document.createElement("li");
      li.className = "head";
      li.innerHTML = `<span class="text">${escape(s)}</span>`;
      list.appendChild(li);
    });
  }

  function renderResults(results) {
    const cnt = { pass: 0, warn: 0, fail: 0, skip: 0 };
    const list = $("results-list");
    list.innerHTML = "";
    results.forEach((r) => {
      cnt[r.status] = (cnt[r.status] || 0) + 1;
      const row = document.createElement("div");
      row.className = "check";
      const actionHtml = r.action
        ? `<button class="secondary" data-action="${escape(r.action.handler)}" style="margin-top:6px;">${escape(r.action.label)}</button>`
        : "";
      // For non-pass results, try to pull the matching lab step(s) so the user
      // sees exactly what they were supposed to do.
      let hintHtml = "";
      if (r.status === "fail" || r.status === "warn") {
        const hints = findRelevantSteps(r, lab);
        if (hints.length) {
          hintHtml =
            `<details class="hint" style="margin-top:8px;">` +
            `<summary style="cursor:pointer;color:var(--accent-2);">📖 What the lab says to do (${hints.length})</summary>` +
            `<ul style="margin:6px 0 0 1.1rem;padding:0;">` +
            hints.map((h) => `<li style="margin:4px 0;">${escape(h)}</li>`).join("") +
            `</ul></details>`;
        }
      }
      row.innerHTML = `
        <span class="pill ${pillClass(r.status)}">${pillLabel(r.status)}</span>
        <span class="name">${escape(r.name)}</span>
        <span></span>
        <span class="msg">${escape(r.msg || "")}${actionHtml ? "<br>" + actionHtml : ""}${hintHtml}</span>
      `;
      list.appendChild(row);
    });
    // Wire up any action buttons
    list.querySelectorAll("button[data-action]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const fn = window.__labCheckBot && window.__labCheckBot[btn.dataset.action];
        if (typeof fn === "function") fn();
      });
    });
    $("cnt-pass").textContent = cnt.pass || 0;
    $("cnt-warn").textContent = cnt.warn || 0;
    $("cnt-fail").textContent = cnt.fail || 0;
    $("cnt-skip").textContent = cnt.skip || 0;
    show("panel-results");
  }

  function pillClass(s) {
    if (s === "pass") return "pass";
    if (s === "warn") return "warn";
    if (s === "fail") return "fail";
    return "info";
  }
  function pillLabel(s) {
    return ({ pass: "PASS", warn: "WARN", fail: "FAIL", skip: "SKIP" }[s]) || s.toUpperCase();
  }

  function escape(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Wire up
  // ──────────────────────────────────────────────────────────────────────────

  document.querySelectorAll(".example").forEach((el) => {
    el.addEventListener("click", () => {
      $("lab-url").value = el.getAttribute("data-url");
      $("btn-load").click();
    });
  });

  $("btn-load").addEventListener("click", async () => {
    showError("");
    const url = $("lab-url").value.trim();
    if (!url) {
      showError("Paste a GitHub URL first.");
      return;
    }
    $("btn-load").disabled = true;
    $("btn-load").textContent = "Loading…";
    try {
      const fetched = await fetchLab(url);
      lab = analyzeLab(fetched);
      persistLab();
      renderLab();
      show("panel-summary");
      show("panel-signin");
      show("panel-run");
    } catch (e) {
      showError(e.message || String(e));
    } finally {
      $("btn-load").disabled = false;
      $("btn-load").textContent = "Load lab";
    }
  });

  $("btn-signin").addEventListener("click", async () => {
    showError("");
    try {
      await signIn();
    } catch (e) {
      showError("Sign-in error: " + (e.message || e));
    }
  });

  $("btn-signout").addEventListener("click", () => signOut());

  $("btn-grant-consent").addEventListener("click", () => {
    showError("");
    // Use /organizations so the user picks (or is already in) any work tenant,
    // and we don't need a tenant id up front. scope=.default lets AAD prompt
    // the GA for *all* statically-configured Graph permissions in one shot.
    const redirectUri =
      location.origin +
      location.pathname.replace(/[^/]*$/, "") +
      "auth-redirect.html";
    const url =
      "https://login.microsoftonline.com/organizations/v2.0/adminconsent" +
      `?client_id=${encodeURIComponent(CLIENT_ID)}` +
      `&scope=${encodeURIComponent("https://graph.microsoft.com/.default")}` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&state=lab-check-bot-graph`;
    showInfo(
      "Opening Microsoft admin-consent page in a new tab. After a Global Admin clicks Accept, come back here and click 'Re-check status' (or sign out and back in if it still shows Not granted)."
    );
    window.open(url, "_blank", "noopener");
  });

  // After admin consent, the existing MSAL refresh token may still not see the
  // newly-consented scopes. The most reliable fix is a full logout + login,
  // which Microsoft's auth server treats as a brand-new session bound to the
  // latest tenant consent state. We set a sessionStorage flag so that after
  // the post-logout redirect completes, we auto-trigger sign-in.
  $("btn-resignin-consent")?.addEventListener("click", async () => {
    showError("");
    showInfo("Signing out and back in to refresh your session with the newly-consented scopes…");
    try {
      sessionStorage.setItem("labCheckBot.autoSignInAfterLogout", "1");
    } catch { /* ignore */ }
    try {
      const inst = await getMsal();
      // postLogoutRedirectUri returns to our own page so the boot code can
      // detect the flag and auto-call loginRedirect.
      await inst.logoutRedirect({ postLogoutRedirectUri: REDIRECT_URI });
    } catch (e) {
      showError("Re-sign in: " + (e.message || e));
    }
  });

  // ── PAT panel wiring ─────────────────────────────────────────────────────
  function refreshPatStatus() {
    const ghStatus = $("pat-github-status");
    if (ghStatus) ghStatus.textContent = getGitHubPat() ? "✅ saved" : "not set";
    const adoStatus = $("pat-ado-status");
    if (adoStatus) {
      const pat = getAdoPat();
      const org = getAdoOrg();
      adoStatus.textContent =
        pat && org ? `✅ saved (org: ${org})` : pat ? "PAT saved · add org" : "not set";
    }
    const ghInput = $("pat-github");
    if (ghInput && !ghInput.value) ghInput.value = getGitHubPat();
    const adoInput = $("pat-ado");
    if (adoInput && !adoInput.value) adoInput.value = getAdoPat();
    const adoOrgInput = $("pat-ado-org");
    if (adoOrgInput && !adoOrgInput.value) adoOrgInput.value = getAdoOrg();
  }
  refreshPatStatus();

  $("btn-save-github-pat")?.addEventListener("click", () => {
    const v = ($("pat-github").value || "").trim();
    setGitHubPat(v);
    refreshPatStatus();
  });
  $("btn-clear-github-pat")?.addEventListener("click", () => {
    setGitHubPat("");
    $("pat-github").value = "";
    refreshPatStatus();
  });
  $("btn-save-ado-pat")?.addEventListener("click", () => {
    setAdoPat(($("pat-ado").value || "").trim());
    setAdoOrg(($("pat-ado-org").value || "").trim());
    refreshPatStatus();
  });
  $("btn-clear-ado-pat")?.addEventListener("click", () => {
    setAdoPat("");
    setAdoOrg("");
    $("pat-ado").value = "";
    $("pat-ado-org").value = "";
    refreshPatStatus();
  });

  $("btn-run").addEventListener("click", async () => {
    if (!lab) {
      showError("Load a lab first.");
      return;
    }
    const account = getActiveAccount();
    if (!account) {
      showError("Sign in first.");
      return;
    }
    showError("");
    $("btn-run").disabled = true;
    $("btn-run").textContent = "Running…";
    try {
      const applicable = STATIC_CHECKS.filter((c) => checkApplies(c, lab));
      const results = [];
      for (const c of applicable) {
        try {
          const out = await c.run();
          results.push({ id: c.id, name: c.name, status: out.status, msg: out.msg });
        } catch (e) {
          results.push({ id: c.id, name: c.name, status: "fail", msg: e.message || String(e) });
        }
      }
      const cs = await runCopilotStudioChecks(lab);
      results.push(...cs);
      const azr = await runAzureChecks(lab);
      results.push(...azr);
      const ghr = await runGitHubChecks(lab);
      results.push(...ghr);
      const ador = await runAdoChecks(lab);
      results.push(...ador);
      renderResults(results);
    } catch (e) {
      showError("Check run failed: " + (e.message || e));
    } finally {
      $("btn-run").disabled = false;
      $("btn-run").textContent = "▶ Run checks";
    }
  });

  // Boot: handle any returning MSAL redirect, update sign-in UI.
  (async () => {
    // Restore any previously-loaded lab BEFORE we await MSAL so a returning
    // redirect doesn't strand the user on an empty page.
    const restored = restoreLab();
    if (restored) {
      lab = restored;
      renderLab();
      show("panel-summary");
      show("panel-signin");
      show("panel-run");
      $("lab-url").value = restored.url || "";
    }
    try {
      await getMsal();
      const account = getActiveAccount();
      if (account) {
        renderAccountInfo(account);
      } else {
        // If we just returned from a logout triggered by the "Re-sign in"
        // button, auto-launch the sign-in flow so the user doesn't have to
        // click again. This is the recommended way to refresh the MSAL
        // refresh token after a tenant admin consent grant.
        let pending = null;
        try { pending = sessionStorage.getItem("labCheckBot.autoSignInAfterLogout"); } catch { /* ignore */ }
        if (pending) {
          try { sessionStorage.removeItem("labCheckBot.autoSignInAfterLogout"); } catch { /* ignore */ }
          try {
            const inst = await getMsal();
            await inst.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
            return;
          } catch (e) {
            showError("Auto re-sign in failed: " + (e.message || e));
          }
        }
      }
      // Clear any stale resume flag from prior versions.
      try { sessionStorage.removeItem("labCheckBot.resumeRun"); } catch { /* ignore */ }
    } catch (e) {
      showError("MSAL init: " + (e.message || e));
    }
  })();

  function renderAccountInfo(account) {
    const userInfo = $("user-info");
    const tenant = (account.username || "").split("@")[1] || "";
    userInfo.innerHTML = `✅ <strong>${escape(account.username || account.name || "Signed in")}</strong>` +
      (tenant ? ` <span class="muted">· tenant: ${escape(tenant)}</span>` : "") +
      ` <a href="#" id="btn-switch-tenant" style="margin-left:8px;">switch tenant</a>`;
    $("btn-signin").classList.add("hidden");
    $("btn-signout").classList.remove("hidden");
    const sw = $("btn-switch-tenant");
    if (sw) {
      sw.addEventListener("click", async (ev) => {
        ev.preventDefault();
        showError("");
        try { await switchTenant(); } catch (e) { showError("Switch tenant: " + (e.message || e)); }
      });
    }
    // Check consent status now that we have an account.
    refreshConsentStatus({ force: false });
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Auto re-check consent when the user comes back from the consent tab
  //
  // auth-redirect.html broadcasts a "consent-granted" message via
  // BroadcastChannel and also writes to localStorage as a fallback. Either
  // signal triggers a forced consent re-check in this tab. We also re-check
  // on window focus (cheap insurance: if a user accepts consent and just
  // alt-tabs back, this catches it even when BroadcastChannel is unavailable).
  // ──────────────────────────────────────────────────────────────────────────

  let lastAutoConsentCheck = 0;
  function autoRecheckConsent(reason) {
    // Throttle to once every 3s to avoid spamming Graph on rapid focus/blur.
    const now = Date.now();
    if (now - lastAutoConsentCheck < 3000) return;
    lastAutoConsentCheck = now;
    if (!getActiveAccount()) return;
    refreshConsentStatus({ forceRefresh: true });
    if (reason) {
      try { console.info("[lab-check-bot] consent re-check triggered:", reason); } catch {}
    }
  }

  // When the consent tab broadcasts success, optimistically mark the card as
  // "consent recorded" so the user isn't confused by a lingering "Not granted".
  // The real verification still happens via Graph — but the MSAL refresh token
  // issued from the pre-consent sign-in often doesn't pick up the new scopes
  // until the user signs out and back in, so we surface that path right here.
  function onConsentBroadcast() {
    setConsentUi(
      "granted",
      "✅ Consent was recorded. If 'Run checks' still fails with 403 errors, click 'Re-sign in' below — admin consent usually needs a fresh sign-in session to take effect."
    );
    // Also try a silent re-check; if it works, the optimistic state stays.
    setTimeout(() => autoRecheckConsent("broadcast (delayed)"), 1500);
  }

  try {
    const bc = new BroadcastChannel("lab-check-bot");
    bc.addEventListener("message", (ev) => {
      if (ev && ev.data && ev.data.type === "lab-check-bot:consent-granted") {
        onConsentBroadcast();
      }
    });
  } catch { /* BroadcastChannel not available; localStorage fallback below */ }

  window.addEventListener("storage", (ev) => {
    if (ev.key === "labCheckBot.consentEvent" && ev.newValue) {
      onConsentBroadcast();
    }
  });

  window.addEventListener("focus", () => autoRecheckConsent("window focus"));

  // ──────────────────────────────────────────────────────────────────────────
  // Consent status indicator
  //
  // Probes a Graph endpoint that requires an admin-consented scope. A 200 means
  // consent is in place for this tenant; a 403 / insufficient privileges means
  // it isn't. We only run this when signed in — without an account we just
  // show "Sign in first" and let the user click the consent button anyway.
  // ──────────────────────────────────────────────────────────────────────────

  function setConsentUi(state, sub) {
    const card = $("consent-card");
    const pill = $("consent-pill");
    const status = $("consent-status");
    if (!card || !pill || !status) return;
    card.classList.remove("granted", "missing");
    pill.classList.remove("pass", "fail", "warn", "info");
    if (state === "granted") {
      card.classList.add("granted");
      pill.classList.add("pass");
      pill.textContent = "✅ Granted";
    } else if (state === "missing") {
      card.classList.add("missing");
      pill.classList.add("fail");
      pill.textContent = "❌ Not granted";
    } else if (state === "checking") {
      pill.classList.add("info");
      pill.textContent = "Checking…";
    } else {
      pill.classList.add("info");
      pill.textContent = "Sign in first";
    }
    if (sub) status.textContent = sub;
  }

  async function refreshConsentStatus({ forceRefresh = false } = {}) {
    const account = getActiveAccount();
    if (!account) {
      setConsentUi(
        "unknown",
        "Sign in above first, then we can verify whether your tenant has already granted consent."
      );
      return;
    }
    setConsentUi("checking", "Checking whether admin consent is already granted for this tenant…");
    try {
      // /applications requires Application.Read.All — fails with 403 if no consent.
      // forceRefresh bypasses the MSAL access-token cache so a freshly-consented
      // tenant doesn't keep returning the old token (which lacked the scopes).
      await graphGet("/applications?$top=1&$select=id", false, { forceRefresh });
      setConsentUi(
        "granted",
        "Admin consent is in place for this tenant. You're ready to run checks."
      );
    } catch (e) {
      const msg = String(e.message || e);
      const looksLikeConsent =
        /403/.test(msg) || /consent/i.test(msg) || /insufficient/i.test(msg) || /interaction_required/i.test(msg);
      setConsentUi(
        "missing",
        looksLikeConsent
          ? "Either admin consent isn't granted yet, or your current sign-in session predates it. Try the buttons below: 'Grant admin consent' (GA only) and then 'Re-sign in' to refresh your session with the newly-consented scopes."
          : `Could not verify consent (${msg}). Try the button below, or 'Re-check status'.`
      );
    }
  }
})();
