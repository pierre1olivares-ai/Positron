# Frontend — Q-Star Issue Manager web part

An SPFx (SharePoint Framework) web part, scaffolded with the Yeoman generator (`@microsoft/generator-sharepoint`, SPFx 1.20, React, Node 18 LTS).

- `prototype/` — the original validated React prototype (`qstar-issue-manager.jsx`) and its clickable demo (`qstar-live.html`). Reference source: the component UI has not yet been ported into the web part below.
- `src/webparts/qstarIssueManager/` — the actual SPFx web part.
  - `components/` — the React component (currently the SPFx boilerplate plus a working **Connection Diagnostics** panel; the prototype's full UI still needs porting in).
  - `models/` — `IIssue.ts` / `ISettings.ts`, typed 1:1 with the prototype's data shapes so porting doesn't require reshaping data.
  - `services/` — two data-layer implementations, both behind the same `IDataService` interface, switchable via the web part's **Data source** property:
    - `BackendApiDataService.ts` — the target architecture: calls the Java backend ([`../backend/service/`](../backend/service/)) over an Azure-AD-secured connection (SPFx's `AadHttpClient`). The backend is a thin gateway to SharePoint, so from here it's just a JSON REST API. Needs the backend actually deployed and an Entra app registration to request tokens against — not usable yet.
    - `SharePointDataService.ts` — talks to the `Q-Star Issues` / `Q-Star Progress Log` lists directly via SharePoint REST (PnPjs), using the signed-in user's own session. No backend deployment or Entra app registration needed — works as soon as the lists are provisioned. This is the practical default (`dataSourceMode: 'sharepoint'`) until the backend has somewhere to run.
    - `MockDataService.ts` — localStorage-backed fallback for UI work in the Workbench before a real list exists.
    - `ConnectionDiagnosticsService.ts` — SharePoint-direct self-test (site access, schema, choice values, indexes, a full write/delete round-trip).
    - `BackendDiagnosticsService.ts` — backend-mode self-test: calls the backend's own `GET /diagnostics`, which proves both legs at once (frontend can reach the backend, and the backend can reach SharePoint).
    - `fieldMap.ts` — single source of truth for SharePoint internal column names; keep in sync with the provisioning scripts and `backend/service/api-contract/contract.yaml`.

Both diagnostics services are wired to the same **Run Connection Test** button in the web part — whichever one matches the active `dataSourceMode`.

## Running locally

```bash
npm install
npx gulp serve   # opens the local Workbench; for a real-tenant test, append --nobrowser
                  # and open https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx yourself
```

To build/package without serving:

```bash
npx gulp bundle --ship
npx gulp package-solution --ship   # produces sharepoint/solution/qstar-issue-manager.sppkg
```

The `.sppkg` is what gets uploaded to the tenant's App Catalog for a real deployment.

## Web part properties

- **Data source**: `sharepoint` (default) or `backend` — picks which data service/diagnostics pair above gets used.
- SharePoint-direct mode: site URL (optional, defaults to the current site), issues/progress list names.
- Backend mode: the backend's base URL (e.g. `https://qstar.time-matters.com/api/v1`) and its Azure AD App ID URI (to request an access token against).

## What's left

1. Port the prototype's UI (`prototype/qstar-issue-manager.jsx`) into `src/webparts/qstarIssueManager/components/QstarIssueManager.tsx`, replacing its `window.storage` calls with the `dataService` prop (already wired through from the web part, works for either data source mode).
2. Remove the Tailwind CDN `<script>` tag present in the prototype — production bundles all styling locally (see the hard constraints in [`../CLAUDE.md`](../CLAUDE.md)). The scaffolded web part already uses SCSS modules, not Tailwind.
3. Wire role resolution (Admin/QM/Owner/Reader) to Entra security groups instead of the prototype's in-app switcher.
4. Before any of the above, either:
   - provision the SharePoint lists (`backend/sharepoint/provisioning/`) and run through [`../backend/sharepoint/connection-test-plan.md`](../backend/sharepoint/connection-test-plan.md) using `dataSourceMode: 'sharepoint'`, or
   - once IT stands up the infra `backend/service/README.md` lists as needed (Postgres, Entra app registrations, Azure DevOps/ACR, Kubernetes), deploy the backend and switch to `dataSourceMode: 'backend'`.
