# Lab Check Bot

Single-page web app that:

1. Takes a GitHub URL to a lab (folder, file, or whole repo).
2. Fetches and parses the markdown.
3. Detects which portals the lab touches and any named items (agents, pools).
4. Signs the user into their Microsoft 365 tenant via MSAL.
5. Runs Microsoft Graph checks to see what's already done.

Everything runs in the browser. No server. Drop the folder on any static host (GitHub Pages is the intended target).

> Note: Copilot Studio agent verification is intentionally Graph-only — we check that the Dataverse service principal is provisioned in the tenant and emit a manual-verify link for each named agent. This avoids needing a Dynamics CRM permission grant in every tenant the tool is used in.

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
5. API permissions (delegated, Microsoft Graph):
   - `User.Read`
   - `Application.Read.All`
   - `Directory.Read.All`
   - `RoleManagement.Read.Directory`
   - `DeviceManagementServiceConfig.Read.All`
   - `DeviceManagementConfiguration.Read.All`
6. (Optional, for Copilot Studio agent name verification) API permissions (delegated, **Dynamics CRM** → `user_impersonation`). If you skip this, the tool falls back to "manual verify" links for each named agent — no consent loop, just less automation.
7. Grant admin consent for the tenant.
8. Copy the Application (client) ID. Either:
   - Edit `app.js` and replace `DEFAULT_CLIENT_ID`, or
   - Append `?clientId=<guid>` to the page URL.

Because the app is multi-tenant, the **first** user from each tenant must consent (Global Admin can grant org-wide).

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

## Files

- `index.html` — UI + styles
- `app.js` — all logic (GitHub fetch, parsing, MSAL, checks, rendering)
- `.github/workflows/pages.yml` — auto-deploy
