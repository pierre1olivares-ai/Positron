# Frontend — Q-Star Issue Manager web part

An SPFx (SharePoint Framework) web part, scaffolded with the Yeoman generator (`@microsoft/generator-sharepoint`, SPFx 1.20, React, Node 18 LTS).

- `prototype/` — the original validated React prototype (`qstar-issue-manager.jsx`) and its clickable demo (`qstar-live.html`). Reference source: the component UI has not yet been ported into the web part below.
- `src/webparts/qstarIssueManager/` — the actual SPFx web part.
  - `components/` — the React component (currently the SPFx boilerplate plus a working **Connection Diagnostics** panel; the prototype's full UI still needs porting in).
  - `models/` — `IIssue.ts` / `ISettings.ts`, typed 1:1 with the prototype's data shapes so porting doesn't require reshaping data.
  - `services/` — the real data layer:
    - `SharePointDataService.ts` — production implementation. Reads/writes the `Q-Star Issues` and `Q-Star Progress Log` lists via SharePoint REST (PnPjs), using the signed-in user's own session. No Entra app registration or Graph admin consent needed — see the rationale in [`../backend/sharepoint/qstar-sharepoint-graph-integration.md`](../backend/sharepoint/qstar-sharepoint-graph-integration.md#4-data-layer--implemented-as-sharepoint-rest-pnpjs-not-graph).
    - `MockDataService.ts` — localStorage-backed fallback for UI work in the Workbench before a real list exists.
    - `ConnectionDiagnosticsService.ts` — live self-test (site access, schema, choice values, indexes, a full write/delete round-trip). Wired to the **Run Connection Test** button in the web part.
    - `fieldMap.ts` — single source of truth for SharePoint internal column names; keep in sync with the provisioning scripts.

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

## What's left

1. Port the prototype's UI (`prototype/qstar-issue-manager.jsx`) into `src/webparts/qstarIssueManager/components/QstarIssueManager.tsx`, replacing its `window.storage` calls with the `dataService` prop (already wired through from the web part).
2. Remove the Tailwind CDN `<script>` tag present in the prototype — production bundles all styling locally (see the hard constraints in [`../CLAUDE.md`](../CLAUDE.md)). The scaffolded web part already uses SCSS modules, not Tailwind.
3. Wire role resolution (Admin/QM/Owner/Reader) to Entra security groups instead of the prototype's in-app switcher.
4. Before any of the above, provision the SharePoint lists (`backend/sharepoint/provisioning/`) and run through [`../backend/sharepoint/connection-test-plan.md`](../backend/sharepoint/connection-test-plan.md) — including the **Run Connection Test** button in this web part — to confirm the tenant connection works before building further on top of it.
