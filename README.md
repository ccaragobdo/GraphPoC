# GraphPoC - FAST Microsoft 365 Staleness Assessment

Browser-only React app (Vite + TypeScript) for a time-boxed Microsoft 365 sampling scan using Microsoft Graph.

This solution is intentionally **not** a full inventory crawler.

## What It Does

- Runs in static browser-only mode with no backend or database
- Supports no-Entra-app operation via pasted delegated Graph token
- Supports optional MSAL sign-in only if a client ID is configured
- Runs a strict time-boxed scan (default 5 minutes)
- Samples SharePoint and optional OneDrive content only
- Applies hard limits for fast execution and partial results
- Computes staleness metrics (>90, >180, >365 days)
- Exports results locally:
  - `file_inventory.csv`
  - `summary_by_site.csv`
  - `summary_by_library.csv`
  - `full_results.json`

## Important Disclaimer

This is a time-based sampling assessment, not a full inventory.

## Default Limits

- Max runtime: 5 minutes
- Max sites: 15
- Max files per site: 400
- Max total files: 5000
- Max folder depth: 2

The scan stops when time or limits are reached and returns partial results.

## No-Entra-App Mode (Default)

If you do not have an Entra app registration, leave `VITE_MSAL_CLIENT_ID` unset.

At runtime, paste a delegated Microsoft Graph bearer token in the app and run the scan.
Required token scopes:

- `User.Read`
- `Files.Read.All`
- `Sites.Read.All`

The app never uploads results and keeps everything in browser memory until export.

## Optional Microsoft Entra App Setup

1. Register a Single Page Application in Microsoft Entra ID.
2. Add delegated Microsoft Graph permissions:
	- `User.Read`
	- `Files.Read.All`
	- `Sites.Read.All`
3. Grant admin consent if required in your tenant.
4. Add redirect URIs:
	- Local: `http://localhost:5173/GraphPoC/`
	- GitHub Pages: `https://<your-org-or-user>.github.io/GraphPoC/`

## Environment Configuration

Create `.env.local` in the repository root:

```bash
VITE_MSAL_CLIENT_ID=<your-app-client-id>
VITE_MSAL_TENANT_ID=common
```

If `VITE_MSAL_CLIENT_ID` is omitted, the app switches automatically to manual-token mode.

## Run Locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## GitHub Pages Deployment

This repository includes [deploy workflow](.github/workflows/deploy.yml) for GitHub Pages.

1. Push to `main`.
2. In GitHub repo settings, enable Pages with **GitHub Actions** source.
3. Ensure app env variables are set before build (or committed in a safe, non-secret config pattern for SPA client IDs).

## Notes

- No backend, database, or external storage is used.
- Results remain in browser memory until user exports/downloads files.
- Scan coverage is limited to content accessible to the signed-in user.
- This is a time-based sampling assessment, not a full inventory.