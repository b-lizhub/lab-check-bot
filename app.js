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
      if (obj && obj.rawText && obj.names) return obj;
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
    // GitHub Pages: <owner>.github.io/<repo>/...  →  treat as the repo root.
    // Org/user pages without a repo segment (<owner>.github.io/) aren't supported.
    if (url.hostname.endsWith(".github.io")) {
      const owner = url.hostname.slice(0, -".github.io".length);
      const parts = url.pathname.split("/").filter(Boolean);
      if (parts.length >= 1) {
        return { kind: "tree", owner, repo: parts[0], ref: "HEAD", path: "" };
      }
      return null;
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
      throw new Error("Could not understand that URL. Use a github.com link.");
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
    { name: "Azure Portal", patterns: [/azure\s*portal/i, /portal\.azure\.com/i, /resource\s*group/i] },
    { name: "Microsoft 365 Admin", patterns: [/microsoft\s*365\s*admin/i, /admin\.microsoft\.com/i] },
    { name: "SharePoint", patterns: [/sharepoint/i] },
    { name: "Teams Admin", patterns: [/teams\s*admin/i] },
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

    return {
      title: fetched.title,
      url: fetched.url,
      rawText: combined,
      steps,
      portals,
      names: { agents, pools },
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

  async function getToken(scopes) {
    const inst = await getMsal();
    const account = getActiveAccount();
    if (!account) throw new Error("Not signed in.");
    const resp = await inst.acquireTokenSilent({ scopes, account });
    return resp.accessToken;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Tenant checks
  // ──────────────────────────────────────────────────────────────────────────

  async function graphGet(path, beta = false) {
    const token = await getToken(["https://graph.microsoft.com/.default"]).catch(
      async () => getToken(GRAPH_SCOPES)
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

  function slug(s) {
    return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
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
    if (lab.names.agents.length === 0 && lab.names.pools.length === 0) {
      named.innerHTML = `<p class="muted">No named agents or pools detected.</p>`;
    } else {
      const parts = [];
      if (lab.names.agents.length) {
        parts.push(
          `<div><strong>Agents:</strong> ${lab.names.agents.map((a) => `<code>${escape(a)}</code>`).join(", ")}</div>`
        );
      }
      if (lab.names.pools.length) {
        parts.push(
          `<div style="margin-top:6px;"><strong>Pools/Groups:</strong> ${lab.names.pools.map((a) => `<code>${escape(a)}</code>`).join(", ")}</div>`
        );
      }
      named.innerHTML = parts.join("");
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
      row.innerHTML = `
        <span class="pill ${pillClass(r.status)}">${pillLabel(r.status)}</span>
        <span class="name">${escape(r.name)}</span>
        <span></span>
        <span class="msg">${escape(r.msg || "")}${actionHtml ? "<br>" + actionHtml : ""}</span>
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
  }
})();
