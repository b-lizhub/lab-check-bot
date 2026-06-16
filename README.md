# Lab Check Bot

Single-page web app that:

1. Takes a GitHub URL to a lab (folder, file, or whole repo).
2. Fetches and parses the markdown.
3. Detects which portals the lab touches (Microsoft 365, Azure, GitHub, Azure DevOps, Copilot Studio, Intune, Entra) and any named items (agents, resources, repos, projects, pipelines).
4. Signs the user into their Microsoft 365 tenant via MSAL.
5. Runs Microsoft Graph + Azure Resource Manager checks, plus optional GitHub and Azure DevOps PAT-based checks, to see what's already done.

Everything runs in the browser. No server. Drop the folder on any static host (GitHub Pages is the intended target).

> Copilot Studio agent verification and Azure resource verification each need one extra one-time admin consent per tenant. GitHub and Azure DevOps checks use a Personal Access Token kept in localStorage. Skip any of them and you'll still get all the other checks plus "what the lab says to do" hints for anything that fails.

## Try it locally

```powershell
cd c:\Users\bsalisbury\repos\lab-check-bot
python -m http.server 5173
# then open http://localhost:5173
```

You can also just double-click `index.html`, but using `http://localhost:5173` makes the redirect URI match what's registered in Entra.

## Required Entra app registration

1. Microsoft Entra admin center → **App registrations** → **New registration**.
2. Name: `Lab Check Bot`.
3. Supported account types: **Accounts in any organizational directory (multi-tenant)**.
4. Redirect URI → platform **Single-page application (SPA)** → add:
   - `http://localhost:5173/`
   - your GitHub Pages URL, e.g. `https://<user>.github.io/lab-check-bot/`
   - and `https://<user>.github.io/lab-check-bot/auth-redirect.html`
5. API permissions (delegated, Microsoft Graph):
   - `User.Read`
   - `Application.Read.All`
   - `Directory.Read.All`
   - `RoleManagement.Read.Directory`
   - `DeviceManagementServiceConfig.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `SecurityIdentitiesSensors.Read.All` (Defender for Identity sensor inventory/health)
   - `SecurityIdentitiesHealth.Read.All` (Defender for Identity health issues)
6. (Optional, for Copilot Studio agent name verification) API permissions (delegated, **Dynamics CRM** → `user_impersonation`).
7. (Optional, for Azure resource verification) API permissions (delegated, **Azure Service Management** → `user_impersonation`). Without this, Azure checks become SKIPs with a one-time "Grant Azure access" button surfaced inline in the results.
8. Grant admin consent for the tenant.
9. Copy the Application (client) ID. Either:
   - Edit `app.js` and replace `DEFAULT_CLIENT_ID`, or
   - Append `?clientId=<guid>` to the page URL.

Because the app is multi-tenant, the **first** user from each tenant must consent (Global Admin can grant org-wide).

### Troubleshooting admin consent

Sign-in itself only requests `User.Read`, so **authenticating never hits the admin-approval wall** — any user can sign in. The admin-only Graph permissions are granted once per tenant via the **Grant admin consent** step.

If the Microsoft page says *"lab checker agent needs permission to access resources in your organization that only an admin can grant"* / *"Need admin approval"*, it means admin consent hasn't been granted for that tenant yet **and** the account that hit the page isn't a Global Administrator. Causes and fixes:

- **The person isn't a Global Admin.** Only a Global Administrator can grant org-wide consent. Have a GA sign in and click **Grant admin consent** once; afterwards every user in that tenant is covered and can sign in normally.
- **Wrong account auto-selected.** The consent page forces an account picker (`prompt=select_account`) — choose a Global Administrator of the tenant you're checking.
- **Consent is pinned to your signed-in tenant.** Sign in to the Lab Check Bot first (step 2) with the target-tenant admin, *then* click Grant admin consent; the request is sent to that exact tenant, not the app's home tenant.
- **Tenant blocks all user consent.** A few tenants require admin approval even for `User.Read`. In that case a GA must approve the app (or run the consent step) before anyone can sign in — this is a tenant policy outside the tool's control.

## Deploy to GitHub Pages

This repo includes `.github/workflows/pages.yml` — push to `main` and it deploys.

In repo settings → **Pages** → set **Source** to **GitHub Actions**.

## What it checks

Static Microsoft Graph checks (only run when the parsed lab text mentions the relevant portal/keywords):

- Windows 365 service principal present
- AVD ARM provider service principal present
- Microsoft Remote Desktop service principal present
- Entra ID P1/P2 license present
- Signed-in user holds a relevant admin role
- Intune enrollment configurations readable

Dynamic Copilot Studio checks (only when the lab mentions Copilot Studio / Power Platform or names agents):

- Dataverse service principal is present and enabled in the tenant (Graph, always runs).
- For each named agent: silently asks for a Dataverse token. If granted, walks every Dataverse environment looking for a bot whose name matches. If not granted, shows a one-time **🔓 Grant Dataverse access** popup button — no redirect, no consent loop.

Azure Resource Manager checks (only when the lab mentions Azure Portal or names resources):

- Lists all enabled subscriptions on the signed-in account.
- For each detected resource group, App Service, Function App, Container App, SQL server, storage account, key vault, or Azure OpenAI / Cognitive Services account: query the matching ARM endpoint across every subscription and PASS/FAIL by name match.
- Requires `Azure Service Management → user_impersonation` consent. If missing, emits one SKIP row with a **🔓 Grant Azure access** button.

GitHub checks (only when the lab mentions GitHub or names repos/workflows):

- Save a PAT once in the "Optional: GitHub & Azure DevOps access tokens" panel (scopes: `repo`, `read:org`, `workflow`).
- Validates the PAT against `/user`, then PASS/FAIL each detected repo (`GET /repos/{owner}/{repo}`) and lists workflows on found repos to match each detected workflow file/name.

Azure DevOps checks (only when the lab mentions ADO or names orgs/projects/pipelines):

- Save a PAT and org name once in the same panel (scopes: Project & Team Read, Code Read, Build Read).
- Lists projects in the org (validates PAT), then PASS/FAIL each detected project and searches each project's build definitions for each detected pipeline name.

## Files

- `index.html` — UI + styles
- `app.js` — all logic (GitHub fetch, parsing, MSAL, Graph + ARM + GitHub + ADO checks, rendering)
- `auth-redirect.html` — landing page after admin-consent redirects
- `.github/workflows/pages.yml` — auto-deploy
