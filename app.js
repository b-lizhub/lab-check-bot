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
    "SecurityIdentitiesHealth.Read.All",
    "SecurityIdentitiesSensors.Read.All",
  ];

  // Scopes requested at *sign-in* only. We deliberately ask for just User.Read
  // (user-consentable in any tenant) so authentication NEVER hits the admin
  // "Need admin approval" wall. The admin-only scopes in GRAPH_SCOPES are
  // granted separately and once-per-tenant via the "Grant admin consent"
  // button; afterwards graphGet acquires them silently with .default.
  const LOGIN_SCOPES = ["User.Read"];


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
        obj.names.entra = obj.names.entra || {};
        obj.names.intune = obj.names.intune || {};
        obj.names.dfi = obj.names.dfi || {};
        obj.names.spSites = obj.names.spSites || [];
        obj.names.openaiDeployments = obj.names.openaiDeployments || [];
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
    { name: "Copilot Studio", patterns: [/copilot\s*studio/i], urlPatterns: [/copilotstudio\.microsoft\.com/i] },
    { name: "Power Platform Admin", patterns: [/power\s*platform\s*admin/i, /power\s*automate/i], urlPatterns: [/admin\.powerplatform/i] },
    { name: "Microsoft Entra", patterns: [/microsoft\s*entra/i, /aad|azure\s*ad\b/i], urlPatterns: [/entra\.microsoft\.com/i] },
    { name: "Microsoft Defender for Identity", patterns: [
        /defender\s*for\s*identity/i, /\bMDI\b/, /azure\s*atp/i,
        /Set-MDIConfiguration/i, /New-MDIConfigurationReport/i, /DefenderForIdentity/i,
        /sensor\s*(?:health|deployment|setup)/i, /honeytoken/i, /gMSA/i,
      ] },
    { name: "Microsoft Sentinel", patterns: [/microsoft\s*sentinel/i, /azure\s*sentinel/i, /sentinel2go/i, /securityinsights/i] },
    { name: "Microsoft Defender XDR", patterns: [/defender\s*xdr/i, /microsoft\s*365\s*defender/i], urlPatterns: [/security\.microsoft\.com/i] },
    { name: "Microsoft Intune", patterns: [/intune/i] },
    { name: "Azure Portal", patterns: [
        /azure\s*portal/i, /resource\s*group/i,
        /app\s*service/i, /container\s*app/i, /function\s*app/i,
        /azure\s*sql/i, /cosmos\s*db/i, /storage\s*account/i,
        /key\s*vault/i, /ai\s*foundry/i, /azure\s*openai/i,
        /log\s*analytics/i, /application\s*insights/i, /\baz\s+\w+/i,
      ], urlPatterns: [/portal\.azure\.com/i] },
    { name: "Microsoft 365 Admin", patterns: [/microsoft\s*365\s*admin/i], urlPatterns: [/admin\.microsoft\.com/i] },
    { name: "SharePoint", patterns: [/sharepoint/i] },
    { name: "Teams Admin", patterns: [/teams\s*admin/i] },
    { name: "GitHub", patterns: [
        /github\s*enterprise/i, /github\s*copilot/i, /github\s*actions/i,
        /github\s*(?:workflow|runner)/i,
        /\bgh\s+(?:auth|repo|workflow|run)\b/i,
      ], urlPatterns: [/\.github\/workflows\//i] },
    { name: "Azure DevOps", patterns: [
        /azure\s*pipelines?\s+(?:run|trigger|create|edit|deploy|build)/i,
        /azure\s*(?:boards?|repos?)\s+(?:create|clone|push|commit|work\s*item)/i,
        /(?:create|configure|set\s*up|import)\s+(?:an?\s+)?azure\s*devops/i,
        /open\s+azure\s*devops/i, /sign\s+in\s+to\s+azure\s*devops/i,
        /navigate\s+to\s+(?:azure\s*devops|dev\.azure\.com)/i,
        ], urlPatterns: [/\bdev\.azure\.com\b/i] },
  ];

  function detectPortals(text, stepText) {
    const found = new Set();
    const operationalText = stepText || text;
    for (const { name, patterns, urlPatterns } of PORTAL_PATTERNS) {
      const matchedOperational = patterns && patterns.some((p) => p.test(operationalText));
      const matchedUrl = urlPatterns && urlPatterns.some((p) => p.test(text));
      if (matchedOperational || matchedUrl) found.add(name);
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

  // ── Entra named entity patterns ────────────────────────────────────────────
  const ENTRA_PATTERNS = {
    groups: [
      /(?:create|add|configure|set\s+up)\s+(?:a\s+)?(?:new\s+)?(?:security\s+|microsoft\s+365\s+|m365\s+|mail-enabled\s+)?group\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /group\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /New-MgGroup\b[^\n]*?-DisplayName\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
    users: [
      /(?:create|add|invite|new)\s+(?:a\s+)?(?:new\s+)?(?:guest\s+)?user\s+(?:account\s+)?(?:with\s+(?:the\s+)?(?:display\s+)?name|named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /New-MgUser\b[^\n]*?-DisplayName\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
    caPolicies: [
      /conditional\s+access\s+policy\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /(?:create|configure|new)\s+(?:a\s+)?(?:new\s+)?(?:conditional\s+access|CA)\s+policy\s+(?:named?|called)?\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
    appRegistrations: [
      /(?:register|create|add|new)\s+(?:a\s+)?(?:new\s+)?(?:an?\s+)?(?:application|app\s+registration|app)\s+(?:with\s+(?:the\s+)?name|named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /New-MgApplication\b[^\n]*?-DisplayName\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
  };

  // ── Intune named entity patterns ───────────────────────────────────────────
  const INTUNE_PATTERNS = {
    compliancePolicies: [
      /(?:create|configure|new|add)\s+(?:a\s+)?(?:new\s+)?(?:device\s+)?compliance\s+policy\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /compliance\s+policy\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
    configProfiles: [
      /(?:create|configure|new|add)\s+(?:a\s+)?(?:new\s+)?(?:device\s+)?configuration\s+profile\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /configuration\s+profile\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
    appProtectionPolicies: [
      /app\s+protection\s+policy\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
      /(?:create|configure|new)\s+(?:a\s+)?(?:new\s+)?app\s+protection\s+policy\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    ],
  };

  // ── SharePoint site patterns ───────────────────────────────────────────────
  const SP_SITE_PATTERNS = [
    /(?:create|configure|new|add)\s+(?:a\s+)?(?:new\s+)?(?:sharepoint\s+)?site\s+(?:named?|called|with\s+(?:the\s+)?name)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
    /sharepoint\s+site\s+(?:named?|called)\s+[`*"']+([^`*"'\n]{2,80})[`*"']+/gi,
  ];

  // ── Azure OpenAI deployment patterns ──────────────────────────────────────
  const OPENAI_DEPLOY_PATTERNS = [
    /(?:create|add|new)\s+(?:a\s+)?(?:model\s+)?deployment\s+(?:named?|called|with\s+(?:the\s+)?name)\s+[`*"']+([A-Za-z0-9][A-Za-z0-9_.-]{1,60})[`*"']+/gi,
    /--deployment-name\s+([A-Za-z0-9][A-Za-z0-9_.-]{1,60})/gi,
    /model\s+deployment\s+(?:named?|called)\s+[`*"']+([A-Za-z0-9][A-Za-z0-9_.-]{1,60})[`*"']+/gi,
    /deployment\s+name[:\s]+[`*"']+([A-Za-z0-9][A-Za-z0-9_.-]{1,60})[`*"']+/gi,
  ];

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
    const stepText = steps.join("\n");
    const portals = detectPortals(combined, stepText);
    const agentsRaw = extractMatches(combined, AGENT_PATTERNS);
    const poolsRaw = extractMatches(combined, POOL_PATTERNS);

    const agents = agentsRaw.filter((a) => !AGENT_STOPWORDS.has(a.toLowerCase()));
    const pools = poolsRaw;

    const azure = extractNamed(combined, AZURE_RESOURCE_PATTERNS);
    const github = extractNamed(combined, GITHUB_PATTERNS);
    const ado = extractNamed(combined, ADO_PATTERNS);
    const entra = extractNamed(combined, ENTRA_PATTERNS);
    const intune = extractNamed(combined, INTUNE_PATTERNS);
    const dfi = extractDfi(combined);
    const spSites = extractMatches(combined, SP_SITE_PATTERNS);
    const openaiDeployments = extractMatches(combined, OPENAI_DEPLOY_PATTERNS);

    return {
      title: fetched.title,
      url: fetched.url,
      rawText: combined,
      steps,
      portals,
      names: { agents, pools, azure, github, ado, entra, intune, dfi, spSites, openaiDeployments },
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
    await inst.loginRedirect({ scopes: LOGIN_SCOPES, prompt: "select_account" });
  }

  async function switchTenant() {
    const inst = await getMsal();
    inst.setActiveAccount(null);
    await inst.loginRedirect({ scopes: LOGIN_SCOPES, prompt: "select_account" });
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
      const err = new Error(`Graph ${resp.status}: ${body || resp.statusText}`);
      err.status = resp.status;
      err.body = body;
      throw err;
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
    try {
      const inst = await getMsal();
      const account = getActiveAccount();
      if (!account) throw new Error("Not signed in");
      // Use acquireTokenPopup so MSAL caches the Dataverse token immediately
      // after consent. That way the next discoverInstances() call (via
      // acquireTokenSilent) succeeds without any extra user action.
      await inst.acquireTokenPopup({
        scopes: [DATAVERSE_DISCOVERY_SCOPE],
        account,
      });
      showInfo("Dataverse access granted — re-running checks…");
      $("btn-run").click();
    } catch (e) {
      showError("Could not grant Dataverse access: " + (e.message || String(e)));
    }
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
      `&prompt=select_account` +
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

    // If no PAT is set, silently skip — the PAT panel explains how to enable these checks.
    if (!pat) return [];

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
      // If PAT is simply expired/invalid (401), return nothing — same as no PAT.
      // The user can clear the stale token in the panel above.
      if (e.status === 401) return [];
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

    // If no PAT/org is set, silently skip — the PAT panel explains how to enable these checks.
    if (!pat || !targetOrg) return [];

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

  // Escape a string for use inside an OData single-quoted literal.
  function oDataStr(s) { return s.replace(/'/g, "''"); }

  // ── Defender for Identity entity extraction ────────────────────────────────
  // Pulls the security-relevant artifacts a Defender for Identity lab expects so
  // checks can validate the *outcome* (and clearly flag what can only be
  // verified on-prem). Generic over the lab text, not hard-coded to one lab.
  function extractDfi(text) {
    const grab = (patterns) => {
      const set = new Set();
      for (const p of patterns) {
        p.lastIndex = 0;
        let m;
        while ((m = p.exec(text)) !== null) {
          const v = (m[1] || "").trim();
          // Drop PowerShell variable-name fragments (e.g. captured from
          // `-Param $gMSA_HostsGroupName`) — these aren't real object names.
          if (v && v.length <= 80 && !v.startsWith("$") && !/^(?:gMSA_|gmsa_)|(?:AccountName|GroupName|HostNames|HostsGroup)$/.test(v)) {
            set.add(v);
          }
        }
      }
      return Array.from(set);
    };

    const gmsaAccounts = grab([
      /\$gMSA_AccountName\s*=\s*['"]([^'"]+)['"]/gi,
      /New-ADServiceAccount\s+-Name\s+['"]?([A-Za-z0-9_$-]{2,40})['"]?/gi,
      /Get-ADServiceAccount\s+-Identity\s+['"]?([A-Za-z0-9_$-]{2,40})['"]?/gi,
      /Directory\s+service\s+accounts?[^\n]*?Account\s+name[^\n]*?\*\*([A-Za-z0-9_-]{2,40})\*\*/gi,
    ]).map((s) => s.replace(/\$$/, ""));

    const gmsaGroups = grab([
      /\$gMSA_HostsGroupName\s*=\s*['"]([^'"]+)['"]/gi,
      /-PrincipalsAllowedToRetrieveManagedPassword\s+\$?([A-Za-z0-9_-]{2,40})/gi,
    ]);

    // Domain controllers / sensor host machines.
    const sensorHosts = [];
    {
      const re = /\$gMSA_HostNames\s*=\s*['"]([^'"]+)['"]/gi;
      let m;
      while ((m = re.exec(text)) !== null) {
        (m[1] || "").split(",").forEach((h) => {
          const v = h.trim();
          if (v) sensorHosts.push(v);
        });
      }
    }

    // MDI audit configurations applied via Set-MDIConfiguration.
    const auditConfigs = grab([
      /Set-MDIConfiguration\s+-Mode\s+\w+\s+-Configuration\s+([A-Za-z]+)/gi,
    ]);

    // Honeytoken accounts: the lab tags a decoy account as a honeytoken.
    const requiresHoneytoken = /honeytoken/i.test(text);
    const honeytokenAccounts = grab([
      /\b(HoneyToken[A-Za-z0-9_]*)\b/g,
    ]);

    return {
      gmsaAccounts,
      gmsaGroups,
      sensorHosts: Array.from(new Set(sensorHosts)),
      auditConfigs,
      requiresHoneytoken,
      honeytokenAccounts,
    };
  }

  function dfiHasAny(dfi) {
    if (!dfi) return false;
    return (
      (dfi.gmsaAccounts || []).length ||
      (dfi.gmsaGroups || []).length ||
      (dfi.sensorHosts || []).length ||
      (dfi.auditConfigs || []).length ||
      dfi.requiresHoneytoken ||
      (dfi.honeytokenAccounts || []).length
    );
  }

  // Build a "not validated" result — the security outcome lives somewhere this
  // browser-based tool can't reach (on-prem AD, local audit policy, a portal
  // setting), or a required permission/scope is missing. We never silently pass
  // these; we say exactly what we couldn't confirm and how to verify by hand.
  function notValidated(o) {
    return {
      id: o.id,
      name: o.name,
      status: "not-validated",
      msg: o.msg || "Could not be validated from the browser.",
      checked: o.checked,
      evidence: o.evidence,
      why: o.why,
      nextStep: o.nextStep,
      labRef: o.labRef,
      action: o.action,
    };
  }

  async function getVmPowerState(vmId) {
    try {
      const data = await armGet(`${vmId}/instanceView?api-version=2023-07-01`);
      const power = (data.statuses || []).find((s) => (s.code || "").startsWith("PowerState/"));
      return power ? power.code.replace("PowerState/", "") : "unknown";
    } catch {
      return "unknown";
    }
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Defender for Identity — deep security validation
  //
  // This validates the *security outcome*, not just that a step happened:
  //   • Sensors actually deployed, reporting, healthy, and mapped to the right
  //     domain controllers (Microsoft Graph security API).
  //   • Defender for Identity's own health issues (audit gaps, stale DCs,
  //     outdated sensors, directory-service-account problems).
  //   • The Azure lab environment (DC + workstation VMs) is actually running.
  //
  // Anything that lives only in on-prem AD (gMSA objects, DC group membership,
  // local audit policy) or in the Defender portal (honeytoken entity tags) is
  // returned as NOT VALIDATED with precise manual verification steps — never a
  // silent pass.
  // ──────────────────────────────────────────────────────────────────────────
  async function runDefenderForIdentityChecks(lab) {
    const dfi = lab.names.dfi || {};
    const isDfiLab =
      (lab.portals || []).includes("Microsoft Defender for Identity") || dfiHasAny(dfi);
    if (!isDfiLab) return [];

    const results = [];

    // ── Fetch MDI sensors + health issues from the Graph security API ──────────
    let sensors = null;
    let sensorsErr = null;
    try {
      const s = await graphGet("/security/identities/sensors", true);
      sensors = s.value || [];
    } catch (e) {
      sensorsErr = e;
    }

    let health = null;
    let healthErr = null;
    try {
      const h = await graphGet("/security/identities/healthIssues", true);
      health = h.value || [];
    } catch (e) {
      healthErr = e;
    }

    const scopeBlocked = (e) => e && (e.status === 401 || e.status === 403);
    const notOnboarded = (e) =>
      e && e.status === 403 && /not onboarded/i.test(e.body || e.message || "");
    const lc = (s) => (s || "").toString().toLowerCase();
    const openIssues = (health || []).filter((i) => lc(i.status) === "open");
    const matchIssues = (re) =>
      openIssues.filter(
        (i) => re.test(i.displayName || "") || re.test(i.issueTypeId || "") || re.test(i.description || "")
      );
    const issueSummary = (issue) => {
      const sev = lc(issue.severity) || "unknown";
      const where = (issue.sensorDNSNames || issue.domainNames || []).join(", ");
      return `${issue.displayName || issue.issueTypeId || "Health issue"} [${sev}${where ? " · " + where : ""}]`;
    };

    // ── 1. Sensor deployment + health (Task 1.3 / Validate the installation) ───
    if (sensors !== null) {
      if (sensors.length === 0) {
        results.push({
          id: "dfi-sensors",
          name: "Defender for Identity sensors deployed",
          status: "missing",
          msg: "No Defender for Identity sensors are reporting in this tenant.",
          checked: "Microsoft Graph security API for installed/reporting MDI sensors.",
          evidence: "GET /security/identities/sensors returned zero sensors.",
          why: "The lab requires an MDI sensor on the domain controller; with no sensor reporting, no detections can fire end-to-end.",
          nextStep:
            "Install the sensor on the DC (download from security.microsoft.com → Settings → Identities → Sensors → Add sensor), paste the access key, and confirm it shows Running.",
          labRef: "Day 1 · Task 1.3: Add a new sensor / Validate the installation.",
        });
      } else {
        const unhealthy = sensors.filter((s) => lc(s.healthStatus) && lc(s.healthStatus) !== "healthy");
        const sensorList = sensors
          .map((s) => `${s.displayName || s.id}${s.version ? " v" + s.version : ""} (${s.healthStatus || "?"})`)
          .join(", ");
        // Map expected DCs from the lab to the reporting sensors.
        const expectedHosts = dfi.sensorHosts || [];
        const reportingNames = sensors.map((s) => lc(s.displayName));
        const missingHosts = expectedHosts.filter(
          (h) => !reportingNames.some((n) => n.includes(lc(h)))
        );
        const sensorHealthIssues = matchIssues(/sensor|stopped communicating|outdated|not reachable|version|service/i);

        let status = "pass";
        const whyParts = [];
        if (unhealthy.length) {
          status = "fail";
          whyParts.push(`${unhealthy.length} sensor(s) report a non-healthy status`);
        }
        if (missingHosts.length) {
          status = "fail";
          whyParts.push(`expected DC(s) not reporting a sensor: ${missingHosts.join(", ")}`);
        }
        if (sensorHealthIssues.length) {
          status = status === "pass" ? "warn" : status;
          whyParts.push(`${sensorHealthIssues.length} open sensor health issue(s)`);
        }
        results.push({
          id: "dfi-sensors",
          name: "Defender for Identity sensors deployed & healthy",
          status,
          msg:
            status === "pass"
              ? `All ${sensors.length} sensor(s) healthy.`
              : "Sensor deployment is not fully healthy.",
          checked: "Sensor inventory, version, health status, and DC mapping via the Graph security API.",
          evidence: `Sensors: ${sensorList}.` + (sensorHealthIssues.length ? ` Open issues: ${sensorHealthIssues.map(issueSummary).join("; ")}.` : ""),
          why:
            status === "pass"
              ? "Sensors are installed, reporting, on a current version, and cover the expected domain controller(s)."
              : whyParts.join("; ") + ".",
          nextStep:
            status === "pass"
              ? "No action — sensors are operational."
              : "In security.microsoft.com → Settings → Identities → Sensors, restart/upgrade unhealthy sensors and install a sensor on any missing DC, then confirm Running + Up to date.",
          labRef: "Day 1 · Task 1.3: Add a new sensor / Validate the installation.",
        });
      }
    } else if (notOnboarded(sensorsErr)) {
      results.push({
        id: "dfi-sensors",
        name: "Defender for Identity sensors deployed & healthy",
        status: "fail",
        msg: "This tenant is not onboarded to Microsoft Defender for Identity.",
        checked: "GET /security/identities/sensors (Microsoft Graph beta).",
        evidence:
          "Graph returned 403: \"Tenant is not onboarded to Microsoft Defender for Identity. After license is purchased, first login to portal and sensor are required.\"",
        why: "MDI onboarding plus at least one reporting sensor is the lab's core deliverable; with no MDI instance, no sensors exist and no detections can fire.",
        nextStep:
          "Ensure a Defender for Identity license is assigned, sign in to security.microsoft.com → Settings → Identities once to initialize the workspace, then install a sensor on the DC and confirm it shows Running.",
        labRef: "Day 1 · Task 1.3: Add a new sensor / Validate the installation.",
      });
    } else if (scopeBlocked(sensorsErr)) {
      results.push(
        notValidated({
          id: "dfi-sensors",
          name: "Defender for Identity sensors deployed & healthy",
          msg: "Sensor inventory could not be read — missing security permission.",
          checked: "GET /security/identities/sensors (Microsoft Graph beta).",
          evidence: `Graph returned ${sensorsErr.status}. The app registration lacks SecurityIdentitiesSensors.Read.All (admin consent), or your account isn't a Security Reader/Admin.`,
          why: "Without this permission the tool cannot confirm sensors are deployed and healthy, so it will not assume they are.",
          nextStep:
            "Have a Global Admin grant admin consent (Settings → grant consent) so SecurityIdentitiesSensors.Read.All is approved, then re-run. To verify manually: security.microsoft.com → Settings → Identities → Sensors → each DC shows Running and Up to date.",
          labRef: "Day 1 · Task 1.3: Add a new sensor / Validate the installation.",
        })
      );
    } else {
      results.push(
        notValidated({
          id: "dfi-sensors",
          name: "Defender for Identity sensors deployed & healthy",
          msg: "Sensor inventory endpoint was not reachable.",
          checked: "GET /security/identities/sensors (Microsoft Graph beta).",
          evidence: sensorsErr ? sensorsErr.message : "No response.",
          why: "The MDI sensors API did not return data (the workspace may not be initialized yet, or the API is unavailable in this tenant).",
          nextStep:
            "Verify manually in security.microsoft.com → Settings → Identities → Sensors that each domain controller sensor shows Running and Up to date.",
          labRef: "Day 1 · Task 1.3: Add a new sensor / Validate the installation.",
        })
      );
    }

    // ── 2. Defender for Identity service health issues ─────────────────────────
    if (health !== null) {
      const nonSensor = openIssues.filter(
        (i) => !/sensor|stopped communicating|outdated|not reachable|version/i.test(i.displayName || "")
      );
      const high = openIssues.filter((i) => lc(i.severity) === "high");
      if (openIssues.length === 0) {
        results.push({
          id: "dfi-health",
          name: "Defender for Identity health issues",
          status: "pass",
          msg: "No open Defender for Identity health issues.",
          checked: "GET /security/identities/healthIssues (open issues across global + sensor scope).",
          evidence: "0 open health issues reported by Defender for Identity.",
          why: "Defender for Identity itself reports the deployment as healthy with no outstanding configuration or connectivity problems.",
          nextStep: "No action.",
          labRef: "Day 1 · Validate the installation.",
        });
      } else {
        results.push({
          id: "dfi-health",
          name: "Defender for Identity health issues",
          status: high.length ? "fail" : "warn",
          msg: `${openIssues.length} open health issue(s)${high.length ? ` (${high.length} high severity)` : ""}.`,
          checked: "GET /security/identities/healthIssues (open issues across global + sensor scope).",
          evidence: openIssues.slice(0, 8).map(issueSummary).join("; "),
          why: "Open health issues mean part of the deployment isn't working as intended; these would suppress or degrade detections.",
          nextStep:
            "Open each issue in security.microsoft.com → Settings → Identities → Health issues and apply the listed remediation, then confirm it clears.",
          labRef: "Day 1 · Validate the installation.",
        });
      }
    } else if (notOnboarded(healthErr)) {
      results.push({
        id: "dfi-health",
        name: "Defender for Identity health issues",
        status: "fail",
        msg: "Health status is unavailable because the tenant isn't onboarded to Defender for Identity.",
        checked: "GET /security/identities/healthIssues (Microsoft Graph beta).",
        evidence:
          "Graph returned 403: tenant is not onboarded to Microsoft Defender for Identity (same root cause as the sensors check).",
        why: "Health data only exists once MDI is onboarded and a sensor is reporting; until then the deployment cannot be confirmed working.",
        nextStep:
          "Complete MDI onboarding and sensor installation (see the sensors check above), then re-run — Health issues should be empty.",
        labRef: "Day 1 · Validate the installation.",
      });
    } else if (scopeBlocked(healthErr)) {
      results.push(
        notValidated({
          id: "dfi-health",
          name: "Defender for Identity health issues",
          msg: "Health issues could not be read — missing security permission.",
          checked: "GET /security/identities/healthIssues (Microsoft Graph beta).",
          evidence: `Graph returned ${healthErr.status}; SecurityIdentitiesHealth.Read.All is not consented or your role lacks access.`,
          why: "Without health data the tool cannot confirm the deployment is actually working, so it will not assume success.",
          nextStep:
            "Grant admin consent for SecurityIdentitiesHealth.Read.All and re-run. Manual check: security.microsoft.com → Settings → Identities → Health issues should be empty.",
          labRef: "Day 1 · Validate the installation.",
        })
      );
    } else {
      results.push(
        notValidated({
          id: "dfi-health",
          name: "Defender for Identity health issues",
          msg: "Health issues endpoint was not reachable.",
          checked: "GET /security/identities/healthIssues (Microsoft Graph beta).",
          evidence: healthErr ? healthErr.message : "No response.",
          why: "The MDI health API did not return data (workspace may still be initializing).",
          nextStep: "Check security.microsoft.com → Settings → Identities → Health issues manually.",
          labRef: "Day 1 · Validate the installation.",
        })
      );
    }

    // ── 3. Audit policy configuration (Task 1.4) ───────────────────────────────
    if (dfi.auditConfigs && dfi.auditConfigs.length) {
      const auditIssues = health !== null ? matchIssues(/audit|advanced audit|ntlm|directory services|configuration container/i) : [];
      if (health !== null && auditIssues.length) {
        results.push({
          id: "dfi-audit",
          name: "Active Directory audit policy configuration",
          status: "fail",
          msg: `${auditIssues.length} audit-related health issue(s) reported by Defender for Identity.`,
          checked: "MDI health issues cross-referenced against the audit policies the lab configures via Set-MDIConfiguration.",
          evidence: auditIssues.map(issueSummary).join("; "),
          why: "Defender for Identity reports that required Windows/AD advanced auditing isn't fully applied, so the related detections won't fire.",
          nextStep:
            "On the DC run New-MDIConfigurationReport -Path C:\\Reports -Mode Domain -OpenHtmlReport, then Set-MDIConfiguration for each failing item until every entry shows Passed.",
          labRef: "Day 1 · Task 1.4: Configure audit policies in AD environment.",
        });
      } else if (health !== null) {
        results.push({
          id: "dfi-audit",
          name: "Active Directory audit policy configuration",
          status: "pass",
          msg: "Defender for Identity reports no outstanding audit-policy gaps.",
          checked: "MDI health issues for audit/NTLM/advanced-auditing gaps, against the lab's expected configurations.",
          evidence: `Expected configs: ${dfi.auditConfigs.join(", ")}. No matching open MDI health issues.`,
          why: "MDI surfaces missing advanced auditing as health issues; none are open, so the required auditing appears applied.",
          nextStep:
            "For full confirmation, run New-MDIConfigurationReport on the DC and verify every item shows Passed.",
          labRef: "Day 1 · Task 1.4: Configure audit policies in AD environment.",
        });
      } else {
        results.push(
          notValidated({
            id: "dfi-audit",
            name: "Active Directory audit policy configuration",
            msg: "Audit policy state lives on the domain controller and can't be read from the browser.",
            checked: "Expected MDI audit configurations parsed from the lab.",
            evidence: `Expected configs: ${dfi.auditConfigs.join(", ")}.`,
            why: "Advanced audit policy and NTLM/object auditing are local AD/GPO settings on the DC, not exposed to Microsoft Graph.",
            nextStep:
              "On the DC run: New-MDIConfigurationReport -Path C:\\Reports -Mode Domain -OpenHtmlReport — every item must show Passed. Fix gaps with Set-MDIConfiguration -Mode Domain -Configuration <Name>.",
            labRef: "Day 1 · Task 1.4: Configure audit policies in AD environment.",
          })
        );
      }
    }

    // ── 4. gMSA / Directory Services Account (Task 1.5 & 1.6) ──────────────────
    if ((dfi.gmsaAccounts && dfi.gmsaAccounts.length) || (dfi.gmsaGroups && dfi.gmsaGroups.length)) {
      const dsaIssues = health !== null ? matchIssues(/directory service|credential|gmsa|service account/i) : [];
      const acct = (dfi.gmsaAccounts || [])[0] || "mdiSvc01";
      const grp = (dfi.gmsaGroups || [])[0] || "mdiSvc01Group";
      const dc = (dfi.sensorHosts || [])[0] || "DC01";
      if (health !== null && dsaIssues.length) {
        results.push({
          id: "dfi-gmsa",
          name: "gMSA / Directory Services Account usable",
          status: "fail",
          msg: "Defender for Identity reports a directory-service-account problem.",
          checked: "MDI health issues for directory service account / credential problems.",
          evidence: dsaIssues.map(issueSummary).join("; "),
          why: "The gMSA exists but MDI cannot use it (wrong group membership, missing 'Log on as a service', or password retrieval blocked), so the sensor can't read AD.",
          nextStep: `On ${dc} run: Get-ADServiceAccount -Identity '${acct}' -Properties PrincipalsAllowedToRetrieveManagedPassword; confirm Get-ADGroupMember '${grp}' contains ${dc}$; verify the 'Log on as a service' GPO and the Directory service account in security.microsoft.com → Settings → Identities.`,
          labRef: "Day 1 · Task 1.5 (Create the gMSA) & Task 1.6 (Configure Group Policy / Directory service accounts).",
        });
      } else {
        results.push(
          notValidated({
            id: "dfi-gmsa",
            name: "gMSA / Directory Services Account usable",
            msg: "gMSA usability is an on-prem AD outcome the browser can't verify.",
            checked: `Expected gMSA "${acct}" and host group "${grp}" parsed from the lab.`,
            evidence:
              health !== null
                ? "No directory-service-account health issues reported by MDI, but that doesn't confirm the gMSA can actually retrieve its password."
                : "MDI health data unavailable.",
            why: "gMSA objects, the host group's membership, KDS root key, deleted-objects-container permissions, and 'Log on as a service' all live in on-prem AD, which Microsoft Graph doesn't expose.",
            nextStep: `On ${dc} run: Test-ADServiceAccount '${acct}' (expect True); Get-ADServiceAccount '${acct}' -Properties PrincipalsAllowedToRetrieveManagedPassword; Get-ADGroupMember '${grp}' must list ${dc}$. Confirm the Directory service account '${acct}' is added under Settings → Identities and is healthy.`,
            labRef: "Day 1 · Task 1.5 (Create the gMSA) & Task 1.6 (Configure Group Policy / Directory service accounts).",
          })
        );
      }
    }

    // ── 5. Honeytoken accounts (Task 4.3) ──────────────────────────────────────
    if (dfi.requiresHoneytoken) {
      const accts = (dfi.honeytokenAccounts || []);
      results.push(
        notValidated({
          id: "dfi-honeytoken",
          name: "Honeytoken account tagged for detection",
          msg: "Honeytoken entity tags live in the Defender portal and aren't exposed to Graph.",
          checked: "Lab requires a honeytoken decoy account to be tagged in Defender for Identity.",
          evidence: accts.length ? `Candidate honeytoken account(s): ${accts.join(", ")}.` : "Honeytoken requirement detected in the lab text.",
          why: "Entity tags (Honeytoken/Sensitive/Exchange) are stored in the Defender for Identity configuration, which the Microsoft Graph delegated APIs used here don't return.",
          nextStep: `In security.microsoft.com → Settings → Identities → Entity tags → Honeytoken, confirm ${accts.length ? "'" + accts[0] + "'" : "the decoy account"} is tagged. Test by signing in as that account and confirming a Honeytoken activity alert fires.`,
          labRef: "Day 2 · Task 4.3: Honeytoken activity.",
        })
      );
    }

    // ── 6. Lab environment health — DC + workstation VMs running (req: usable) ──
    if ((dfi.sensorHosts && dfi.sensorHosts.length) || /\bWIN\d+\b|\bDC0?1\b/.test(lab.rawText)) {
      let subs = null;
      try {
        subs = await listAllSubscriptions();
      } catch {
        subs = null;
      }
      if (subs && subs.length) {
        const vms = [];
        for (const sub of subs) {
          try {
            const items = await listResourcesOfType(sub.id, "Microsoft.Compute/virtualMachines");
            for (const v of items) vms.push(v);
          } catch {
            /* ignore one sub */
          }
        }
        const labVms = vms.filter((v) => /^DC0?\d|^WIN\d+/i.test(v.name));
        if (labVms.length === 0) {
          results.push({
            id: "dfi-env",
            name: "Lab environment (DC + workstations) deployed",
            status: "missing",
            msg: "No domain controller / workstation VMs found in accessible subscriptions.",
            checked: "Azure Resource Manager for the lab's DC and WIN* virtual machines.",
            evidence: vms.length ? `Found ${vms.length} VM(s) but none named like DC*/WIN*.` : "No VMs found.",
            why: "Without the DC and workstation VMs, the sensor has nothing to monitor and the attack simulations can't run.",
            nextStep: "Deploy the lab environment (Microsoft-Sentinel2Go template) and confirm DC01/WIN5/WIN6 exist and are running.",
            labRef: "Day 1 · Task 1.1: Configure the Azure environment.",
          });
        } else {
          const states = [];
          for (const v of labVms) states.push({ name: v.name, state: await getVmPowerState(v.id) });
          const stopped = states.filter((s) => s.state !== "running");
          results.push({
            id: "dfi-env",
            name: "Lab environment (DC + workstations) running",
            status: stopped.length ? "warn" : "pass",
            msg: stopped.length ? `${stopped.length} lab VM(s) not running.` : `All ${states.length} lab VM(s) running.`,
            checked: "Power state of the DC and workstation VMs via Azure Resource Manager.",
            evidence: states.map((s) => `${s.name}: ${s.state}`).join(", "),
            why: stopped.length
              ? "Stopped VMs mean the sensor, DC, or attack workstations are offline, so detections can't be exercised end-to-end."
              : "The domain controller and workstations are powered on and usable.",
            nextStep: stopped.length ? "Start the stopped VMs in the Azure portal before running the lab exercises." : "No action.",
            labRef: "Day 1 · Task 1.1: Configure the Azure environment.",
          });
        }
      }
      // If ARM isn't consented, runAzureChecks already emits the grant-access
      // action, so we don't duplicate a not-validated row here.
    }

    return results;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Entra named-entity checks (groups, users, CA policies, app registrations)
  // ──────────────────────────────────────────────────────────────────────────

  async function runEntraNamedChecks(lab) {
    const entra = lab.names.entra || {};
    const hasAny = Object.values(entra).some((arr) => arr && arr.length);
    if (!hasAny) return [];

    const results = [];

    // Groups — existence is not enough: a security group with no members can't
    // actually grant access, so an empty group is only a warning.
    for (const name of entra.groups || []) {
      try {
        const d = await graphGet(`/groups?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName,groupTypes,securityEnabled`);
        const hit = (d.value || [])[0];
        if (!hit) {
          results.push({
            id: `entra-group-${slug(name)}`,
            name: `Entra group "${name}"`,
            status: "fail",
            msg: "Group not found in directory.",
            checked: "Directory for a group with this display name.",
            evidence: "No group matched the display name.",
            why: "The group the lab asks you to create does not exist, so nothing downstream that targets it can work.",
            nextStep: "Create the group in Entra ID → Groups exactly as the lab specifies.",
          });
          continue;
        }
        let memberCount = null;
        try {
          const mem = await graphGet(`/groups/${hit.id}/members?$select=id&$top=20`);
          memberCount = (mem.value || []).length;
        } catch { /* membership read may be blocked; leave null */ }
        const isM365 = (hit.groupTypes || []).includes("Unified");
        const empty = memberCount === 0;
        results.push({
          id: `entra-group-${slug(name)}`,
          name: `Entra group "${name}"`,
          status: empty ? "warn" : "pass",
          msg: empty ? "Found, but has no members." : "Found and populated.",
          checked: "Group existence, type, and membership.",
          evidence: `${isM365 ? "Microsoft 365" : "Security"} group · ${memberCount === null ? "membership not readable" : memberCount + (memberCount === 20 ? "+" : "") + " member(s)"}.`,
          why: empty
            ? "The group exists but is empty, so any access or policy that relies on its membership won't take effect."
            : "The group exists and contains members, so membership-based access/policy can apply.",
          nextStep: empty ? "Add the expected members to the group as the lab describes." : "No action.",
        });
      } catch (e) {
        results.push({ id: `entra-group-${slug(name)}`, name: `Entra group "${name}"`, status: "warn", msg: `Could not check: ${e.message}`, why: "The directory query failed.", nextStep: "Re-run after confirming Directory.Read.All consent." });
      }
    }

    // Users — existence + whether the account is actually enabled/usable.
    for (const name of entra.users || []) {
      try {
        const token = await getToken(["https://graph.microsoft.com/.default"]);
        const resp = await fetch(
          `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName,userPrincipalName,accountEnabled&$count=true`,
          { headers: { Authorization: "Bearer " + token, ConsistencyLevel: "eventual" } }
        );
        const d = resp.ok ? await resp.json() : { value: [] };
        const hit = (d.value || [])[0];
        if (!hit) {
          results.push({
            id: `entra-user-${slug(name)}`,
            name: `Entra user "${name}"`,
            status: "fail",
            msg: "User not found in directory.",
            checked: "Directory for a user with this display name.",
            evidence: "No matching user.",
            why: "The account the lab asks you to create is missing.",
            nextStep: "Create the user in Entra ID → Users as the lab specifies.",
          });
          continue;
        }
        results.push({
          id: `entra-user-${slug(name)}`,
          name: `Entra user "${name}"`,
          status: hit.accountEnabled === false ? "warn" : "pass",
          msg: hit.accountEnabled === false ? "Found, but the account is disabled." : `Found · ${hit.userPrincipalName}`,
          checked: "User existence and account-enabled state.",
          evidence: `${hit.userPrincipalName} · accountEnabled: ${hit.accountEnabled}`,
          why: hit.accountEnabled === false ? "A disabled account can't sign in, so it can't be used in the lab scenario." : "The account exists and is enabled.",
          nextStep: hit.accountEnabled === false ? "Enable the account in Entra ID → Users." : "No action.",
        });
      } catch (e) {
        results.push({ id: `entra-user-${slug(name)}`, name: `Entra user "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    // Conditional Access — must be enabled AND have conditions + grant controls
    // to actually enforce anything. Report-only / disabled is not a pass.
    for (const name of entra.caPolicies || []) {
      try {
        const d = await graphGet(`/identity/conditionalAccess/policies?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName,state,conditions,grantControls`);
        const hit = (d.value || [])[0];
        if (!hit) {
          results.push({
            id: `entra-ca-${slug(name)}`,
            name: `Conditional Access policy "${name}"`,
            status: "fail",
            msg: "Policy not found.",
            checked: "Conditional Access policies by display name.",
            evidence: "No matching policy.",
            why: "The policy the lab asks you to create doesn't exist, so it enforces nothing.",
            nextStep: "Create it in Entra ID → Protection → Conditional Access.",
          });
          continue;
        }
        const hasControls = !!(hit.grantControls && ((hit.grantControls.builtInControls || []).length || (hit.grantControls.customAuthenticationFactors || []).length || hit.grantControls.termsOfUse));
        const hasAssignments = !!(hit.conditions && hit.conditions.users && ((hit.conditions.users.includeUsers || []).length || (hit.conditions.users.includeGroups || []).length || (hit.conditions.users.includeRoles || []).length));
        let status = "warn";
        let why = "";
        if (hit.state === "enabled" && hasControls && hasAssignments) {
          status = "pass";
          why = "The policy is enabled and has both target assignments and grant controls, so it actively enforces the intended security outcome.";
        } else if (hit.state !== "enabled") {
          status = "warn";
          why = `The policy exists but is in '${hit.state}' state, so it is not enforcing anything yet.`;
        } else {
          status = "warn";
          why = "The policy is enabled but is missing user/group assignments or grant controls, so it has no real effect.";
        }
        results.push({
          id: `entra-ca-${slug(name)}`,
          name: `Conditional Access policy "${name}"`,
          status,
          msg: `Found · state: ${hit.state}`,
          checked: "Policy state, target assignments, and grant controls.",
          evidence: `state=${hit.state}; assignments=${hasAssignments ? "yes" : "none"}; grantControls=${hasControls ? "yes" : "none"}.`,
          why,
          nextStep: status === "pass" ? "No action." : "Set the policy to 'On', assign target users/groups, and configure grant controls.",
        });
      } catch (e) {
        results.push({ id: `entra-ca-${slug(name)}`, name: `CA policy "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    // App registrations — existence + a usable enterprise app (service
    // principal) + credentials/permissions, since an app with none of those
    // can't authenticate or call anything.
    for (const name of entra.appRegistrations || []) {
      try {
        const d = await graphGet(`/applications?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName,appId,passwordCredentials,keyCredentials,requiredResourceAccess`);
        const hit = (d.value || [])[0];
        if (!hit) {
          results.push({
            id: `entra-app-${slug(name)}`,
            name: `App registration "${name}"`,
            status: "fail",
            msg: "App registration not found.",
            checked: "App registrations by display name.",
            evidence: "No matching application.",
            why: "The app the lab asks you to register doesn't exist.",
            nextStep: "Register it in Entra ID → App registrations.",
          });
          continue;
        }
        const now = Date.now();
        const validSecrets = (hit.passwordCredentials || []).filter((c) => !c.endDateTime || new Date(c.endDateTime).getTime() > now).length;
        const validCerts = (hit.keyCredentials || []).filter((c) => !c.endDateTime || new Date(c.endDateTime).getTime() > now).length;
        const permCount = (hit.requiredResourceAccess || []).reduce((n, r) => n + (r.resourceAccess || []).length, 0);
        let spExists = null;
        try {
          const sp = await graphGet(`/servicePrincipals?$filter=appId eq '${oDataStr(hit.appId)}'&$select=id,accountEnabled`);
          spExists = (sp.value || [])[0] || null;
        } catch { /* ignore */ }
        const issues = [];
        if (spExists === null) issues.push("enterprise app (service principal) not provisioned");
        if (validSecrets + validCerts === 0) issues.push("no valid client secret or certificate");
        const status = issues.length ? "warn" : "pass";
        results.push({
          id: `entra-app-${slug(name)}`,
          name: `App registration "${name}"`,
          status,
          msg: status === "pass" ? "Found and usable." : "Found, but not fully usable.",
          checked: "App existence, service principal, credentials, and API permissions.",
          evidence: `client ID ${hit.appId} · secrets:${validSecrets} certs:${validCerts} · API permissions:${permCount} · SP:${spExists ? "present" : "missing"}.`,
          why: status === "pass"
            ? "The app exists, has a service principal, and holds valid credentials, so it can authenticate and call its configured APIs."
            : `The app exists but ${issues.join(" and ")}, so it can't authenticate or be granted access end-to-end.`,
          nextStep: status === "pass" ? "No action." : "Add a client secret/certificate and ensure the enterprise application (service principal) and required API permissions (with admin consent) are in place.",
        });
      } catch (e) {
        results.push({ id: `entra-app-${slug(name)}`, name: `App registration "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    return results;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Intune named-entity checks (compliance policies, config profiles, app protection)
  // ──────────────────────────────────────────────────────────────────────────

  async function runIntuneNamedChecks(lab) {
    const intune = lab.names.intune || {};
    const hasAny = Object.values(intune).some((arr) => arr && arr.length);
    if (!hasAny) return [];

    const results = [];

    for (const name of intune.compliancePolicies || []) {
      try {
        const d = await graphGet(`/deviceManagement/deviceCompliancePolicies?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName`, true);
        const hit = (d.value || [])[0];
        results.push({
          id: `intune-comp-${slug(name)}`,
          name: `Intune compliance policy "${name}"`,
          status: hit ? "pass" : "fail",
          msg: hit ? `Found` : "Compliance policy not found. Create it in Intune → Devices → Compliance policies.",
        });
      } catch (e) {
        results.push({ id: `intune-comp-${slug(name)}`, name: `Intune compliance policy "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    for (const name of intune.configProfiles || []) {
      try {
        const d = await graphGet(`/deviceManagement/deviceConfigurations?$filter=displayName eq '${oDataStr(name)}'&$select=id,displayName`, true);
        const hit = (d.value || [])[0];
        results.push({
          id: `intune-cfg-${slug(name)}`,
          name: `Intune config profile "${name}"`,
          status: hit ? "pass" : "fail",
          msg: hit ? `Found` : "Configuration profile not found. Create it in Intune → Devices → Configuration.",
        });
      } catch (e) {
        results.push({ id: `intune-cfg-${slug(name)}`, name: `Intune config profile "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    for (const name of intune.appProtectionPolicies || []) {
      try {
        const d = await graphGet(`/deviceAppManagement/managedAppPolicies?$select=id,displayName`, true);
        const hit = (d.value || []).find((p) => (p.displayName || "").toLowerCase() === name.toLowerCase());
        results.push({
          id: `intune-app-${slug(name)}`,
          name: `Intune app protection policy "${name}"`,
          status: hit ? "pass" : "fail",
          msg: hit ? `Found` : "App protection policy not found. Create it in Intune → Apps → App protection policies.",
        });
      } catch (e) {
        results.push({ id: `intune-app-${slug(name)}`, name: `Intune app protection policy "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }

    return results;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // SharePoint site checks
  // ──────────────────────────────────────────────────────────────────────────

  async function runSharePointChecks(lab) {
    const sites = lab.names.spSites || [];
    if (!sites.length) return [];

    const results = [];
    for (const name of sites) {
      try {
        const d = await graphGet(`/sites?search=${encodeURIComponent(name)}&$select=id,displayName,name,webUrl`);
        const hit = (d.value || []).find(
          (s) =>
            (s.displayName || "").toLowerCase() === name.toLowerCase() ||
            (s.name || "").toLowerCase() === name.toLowerCase()
        );
        results.push({
          id: `sp-site-${slug(name)}`,
          name: `SharePoint site "${name}"`,
          status: hit ? "pass" : "fail",
          msg: hit ? `Found · ${hit.webUrl}` : "Site not found. Create it in SharePoint admin or via a lab step.",
        });
      } catch (e) {
        results.push({ id: `sp-site-${slug(name)}`, name: `SharePoint site "${name}"`, status: "warn", msg: `Could not check: ${e.message}` });
      }
    }
    return results;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Azure OpenAI model deployment checks
  // ──────────────────────────────────────────────────────────────────────────

  async function runOpenAiDeploymentChecks(lab) {
    const deploymentNames = lab.names.openaiDeployments || [];
    const openaiResources = (lab.names.azure || {}).openAiAccounts || [];
    if (!deploymentNames.length || !openaiResources.length) return [];

    // Reuse ARM token — if not yet acquired, this will fail gracefully.
    let subs;
    try { subs = await listAllSubscriptions(); } catch { return []; }
    if (!subs || !subs.length) return [];

    // Find the OpenAI resources by name across subscriptions.
    const foundResources = [];
    for (const sub of subs) {
      try {
        const items = await listResourcesOfType(sub.id, "Microsoft.CognitiveServices/accounts");
        for (const r of items) {
          if (openaiResources.some((n) => n.toLowerCase() === r.name.toLowerCase())) {
            foundResources.push({ id: r.id, name: r.name });
          }
        }
      } catch { /* ignore */ }
    }
    if (!foundResources.length) return [];

    const results = [];
    for (const dep of deploymentNames) {
      let found = null;
      for (const res of foundResources) {
        try {
          const d = await armGet(`${res.id}/deployments?api-version=2023-05-01`);
          const hit = (d.value || []).find((x) => (x.name || "").toLowerCase() === dep.toLowerCase());
          if (hit) { found = { resource: res.name, model: hit.properties?.model?.name || "?" }; break; }
        } catch { /* ignore */ }
      }
      results.push({
        id: `openai-dep-${slug(dep)}`,
        name: `Azure OpenAI deployment "${dep}"`,
        status: found ? "pass" : "fail",
        msg: found
          ? `Found in "${found.resource}" · model: ${found.model}`
          : `Not found in ${foundResources.map((r) => `"${r.name}"`).join(", ")}. Deploy the model in Azure AI Foundry or the Azure portal.`,
      });
    }
    return results;
  }
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
    const entra = lab.names.entra || {};
    const intune = lab.names.intune || {};
    const dfi = lab.names.dfi || {};
    const hasAnyName =
      lab.names.agents.length ||
      lab.names.pools.length ||
      (lab.names.spSites || []).length ||
      (lab.names.openaiDeployments || []).length ||
      Object.values(az).some((v) => v && v.length) ||
      Object.values(gh).some((v) => v && v.length) ||
      Object.values(ado).some((v) => v && v.length) ||
      Object.values(entra).some((v) => v && v.length) ||
      Object.values(intune).some((v) => v && v.length) ||
      dfiHasAny(dfi);
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
      // Entra
      parts.push(row("Entra groups", entra.groups));
      parts.push(row("Entra users", entra.users));
      parts.push(row("Conditional Access policies", entra.caPolicies));
      parts.push(row("App registrations", entra.appRegistrations));
      // Intune
      parts.push(row("Intune compliance policies", intune.compliancePolicies));
      parts.push(row("Intune config profiles", intune.configProfiles));
      parts.push(row("Intune app protection policies", intune.appProtectionPolicies));
      // Defender for Identity
      parts.push(row("MDI gMSA service accounts", dfi.gmsaAccounts));
      parts.push(row("MDI gMSA host groups", dfi.gmsaGroups));
      parts.push(row("MDI sensor hosts (DCs)", dfi.sensorHosts));
      parts.push(row("MDI audit configurations", dfi.auditConfigs));
      parts.push(row("Honeytoken accounts", dfi.honeytokenAccounts));
      // SharePoint
      if ((lab.names.spSites || []).length) parts.push(row("SharePoint sites", lab.names.spSites));
      // Azure
      parts.push(row("Azure resource groups", az.resourceGroups));
      parts.push(row("Azure App Services / Web Apps", az.appServices));
      parts.push(row("Azure Container Apps", az.containerApps));
      parts.push(row("Azure Function Apps", az.functionApps));
      parts.push(row("Azure SQL servers", az.sqlServers));
      parts.push(row("Azure Storage accounts", az.storageAccounts));
      parts.push(row("Azure Key Vaults", az.keyVaults));
      parts.push(row("Azure OpenAI / AI Foundry", az.openAiAccounts));
      if ((lab.names.openaiDeployments || []).length) parts.push(row("Azure OpenAI model deployments", lab.names.openaiDeployments));
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

    updateTokenPanel(lab);
  }

  function updateTokenPanel(lab) {
    const portals = lab.portals || [];
    const gh = lab.names.github || {};
    const ado = lab.names.ado || {};
    const az = lab.names.azure || {};

    const needsGitHub = portals.includes("GitHub");

    const needsAdo = portals.includes("Azure DevOps");

    const needsAzure =
      portals.includes("Azure Portal") ||
      Object.values(az).some((v) => v && v.length > 0);

    // Update per-section badges
    const setBadge = (id, needed) => {
      const el = $(id);
      if (!el) return;
      if (needed) {
        el.textContent = "Required";
        el.style.background = "rgba(255,160,60,0.18)";
        el.style.borderColor = "rgba(242,140,30,0.6)";
        el.style.color = "var(--text)";
      } else {
        el.textContent = "Not applicable";
        el.style.background = "";
        el.style.borderColor = "";
        el.style.color = "var(--muted)";
      }
    };
    setBadge("pat-github-badge", needsGitHub);
    setBadge("pat-ado-badge", needsAdo);

    // Update panel-level summary badge + auto-open/close
    const panelBadge = $("pat-panel-badge");
    const panel = $("pat-panel");
    if (needsGitHub || needsAdo) {
      const parts = [];
      if (needsGitHub) parts.push("GitHub");
      if (needsAdo) parts.push("Azure DevOps");
      panelBadge.textContent = `Required — ${parts.join(" & ")}`;
      panelBadge.style.background = "rgba(255,160,60,0.18)";
      panelBadge.style.borderColor = "rgba(242,140,30,0.6)";
      panelBadge.style.color = "var(--text)";
      panel.open = true;
    } else {
      panelBadge.textContent = "Not applicable for this lab";
      panelBadge.style.background = "";
      panelBadge.style.borderColor = "";
      panelBadge.style.color = "var(--muted)";
      panel.open = false;
    }

    // Dim not-applicable sections
    const ghSection = $("pat-github-section");
    const adoSection = $("pat-ado-section");
    if (ghSection) ghSection.style.opacity = needsGitHub ? "" : "0.45";
    if (adoSection) adoSection.style.opacity = needsAdo ? "" : "0.45";
  }

  function renderResults(results) {
    const cnt = { pass: 0, warn: 0, fail: 0, skip: 0, "not-validated": 0, missing: 0 };
    const list = $("results-list");
    list.innerHTML = "";
    results.forEach((r) => {
      cnt[r.status] = (cnt[r.status] || 0) + 1;
      const row = document.createElement("div");
      row.className = "check";
      const actionHtml = r.action
        ? `<button class="secondary" data-action="${escape(r.action.handler)}" style="margin-top:6px;">${escape(r.action.label)}</button>`
        : "";

      // Explainable review block: what was checked, evidence, why, next step,
      // and the lab step it maps to. Each field is optional; legacy checks that
      // only set `msg` still render fine.
      const reviewLines = [];
      if (r.checked) reviewLines.push(`<div class="ln"><span class="k">Checked</span>${escape(r.checked)}</div>`);
      if (r.evidence) reviewLines.push(`<div class="ln"><span class="k">Evidence</span>${escape(r.evidence)}</div>`);
      if (r.why) reviewLines.push(`<div class="ln why"><span class="k">Why</span>${escape(r.why)}</div>`);
      if (r.nextStep) reviewLines.push(`<div class="ln next"><span class="k">Next</span>${escape(r.nextStep)}</div>`);

      // Lab step mapping: prefer an explicit labRef, else fall back to the
      // heuristic matcher for fail/warn/missing/not-validated rows.
      let refHtml = "";
      if (r.labRef) {
        refHtml = `<div class="ln ref"><span class="k">Lab step</span>${escape(r.labRef)}</div>`;
      } else if (["fail", "warn", "missing", "not-validated"].includes(r.status)) {
        const hints = findRelevantSteps(r, lab);
        if (hints.length) {
          refHtml =
            `<details class="hint" style="margin-top:4px;">` +
            `<summary style="cursor:pointer;color:var(--accent-2);font-size:0.82rem;">📖 What the lab says to do (${hints.length})</summary>` +
            `<ul style="margin:6px 0 0 1.1rem;padding:0;">` +
            hints.map((h) => `<li style="margin:4px 0;font-size:0.82rem;">${escape(h)}</li>`).join("") +
            `</ul></details>`;
        }
      }
      const reviewHtml = (reviewLines.length || refHtml || actionHtml)
        ? `<div class="review">${reviewLines.join("")}${refHtml}${actionHtml}</div>`
        : "";

      row.innerHTML = `
        <span class="pill ${pillClass(r.status)}">${pillLabel(r.status)}</span>
        <span class="name">${escape(r.name)}</span>
        <span></span>
        <span class="msg">${escape(r.msg || "")}</span>
        ${reviewHtml}
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
    $("cnt-missing").textContent = cnt.missing || 0;
    $("cnt-notvalidated").textContent = cnt["not-validated"] || 0;
    $("cnt-skip").textContent = cnt.skip || 0;
    show("panel-results");
  }

  function pillClass(s) {
    if (s === "pass") return "pass";
    if (s === "warn") return "warn";
    if (s === "fail") return "fail";
    if (s === "missing") return "missing";
    if (s === "not-validated") return "notvalidated";
    return "info";
  }
  function pillLabel(s) {
    return ({
      pass: "PASS",
      warn: "WARN",
      fail: "FAIL",
      skip: "SKIP",
      missing: "EXPECTED MISSING",
      "not-validated": "NOT VALIDATED",
    }[s]) || s.toUpperCase();
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

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.getAttribute("data-tab");
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + tab).classList.add("active");
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
    // Pin the consent request to the tenant the user is actually signed in to
    // (the tenant being checked), so the grant lands in the right place and we
    // never silently reuse a cached session for the app's home ("lab checker")
    // tenant. Fall back to /organizations only if we can't resolve a tenant.
    // prompt=select_account forces the Global Admin to pick the correct account
    // instead of AAD auto-using whatever session happens to be cached — the root
    // cause of testers seeing "this app needs admin approval" for the wrong tenant.
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
      `&scope=${encodeURIComponent("https://graph.microsoft.com/.default")}` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&prompt=select_account` +
      `&state=lab-check-bot-graph`;
    showInfo(
      "Opening the Microsoft admin-consent page in a new tab. You must pick a " +
      "Global Administrator account for the tenant you're checking. If you see " +
      "\"needs admin approval,\" the account you chose isn't a Global Admin in " +
      "that tenant — pick a different account. After consent, come back and click " +
      "'Re-check status' (or sign out and back in if it still shows Not granted)."
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
      const entraR = await runEntraNamedChecks(lab);
      results.push(...entraR);
      const dfiR = await runDefenderForIdentityChecks(lab);
      results.push(...dfiR);
      const intuneR = await runIntuneNamedChecks(lab);
      results.push(...intuneR);
      const spR = await runSharePointChecks(lab);
      results.push(...spR);
      const azr = await runAzureChecks(lab);
      results.push(...azr);
      const openaiR = await runOpenAiDeploymentChecks(lab);
      results.push(...openaiR);
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
            await inst.loginRedirect({ scopes: LOGIN_SCOPES, prompt: "select_account" });
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
    maybeWarnCorpTenant(account, tenant);
    // Check consent status now that we have an account.
    refreshConsentStatus({ force: false });
  }

  // The Microsoft corporate tenant. Signing in here is the #1 tester mistake:
  // it's almost never the tenant the lab was deployed to, the tester isn't a
  // Global Admin of it, and its consent policy will block the app anyway.
  const MICROSOFT_CORP_TENANT_ID = "72f988bf-86f1-41af-91ab-2d7cd011db47";

  function maybeWarnCorpTenant(account, tenantDomain) {
    const el = $("tenant-warning");
    if (!el) return;
    const tid =
      (account && (account.tenantId || (account.idTokenClaims && account.idTokenClaims.tid))) || "";
    const domain = (tenantDomain || "").toLowerCase();
    const isCorp =
      tid === MICROSOFT_CORP_TENANT_ID ||
      domain === "microsoft.com" ||
      domain === "microsoft.onmicrosoft.com";
    if (!isCorp) {
      el.classList.add("hidden");
      el.innerHTML = "";
      return;
    }
    el.innerHTML =
      `<span class="tenant-warning-title">⚠️ You're signed into the Microsoft corporate tenant</span>` +
      `An <strong>@microsoft.com</strong> account checks the Microsoft production tenant — <strong>not</strong> the demo/CDX tenant where your lab was deployed. ` +
      `The lab's resources won't be there, you almost certainly aren't a Global Admin of it, and corporate consent policy will block the app. ` +
      `Click <strong>switch tenant</strong> above and sign in with the lab's demo-tenant admin account (for example <code>admin@&lt;tenant&gt;.onmicrosoft.com</code>).`;
    el.classList.remove("hidden");
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
          ? "Either admin consent isn't granted yet, or your current sign-in session predates it. Click 'Grant admin consent' below and pick a Global Administrator account for THIS tenant — if the Microsoft page says \"needs admin approval,\" the account you picked isn't a Global Admin. Then click 'Re-sign in' to refresh your session with the newly-consented scopes."
          : `Could not verify consent (${msg}). Try the button below, or 'Re-check status'.`
      );
    }
  }
})();
