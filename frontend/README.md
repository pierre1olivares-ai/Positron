# Frontend — Q-Star Issue Manager web part

Production SPFx 1.20 / React 17 web part for the validated Q-Star Issue Manager.

## Implemented

- The complete validated prototype UI is ported to `components/QstarPrototype.tsx` and uses `IDataService`; production no longer uses `window.storage`.
- `SharePointDataService.ts` reads and writes the `Q-Star Issues`, `Q-Star Progress Log`, and `Q-Star Config` lists through same-site SharePoint REST/PnPjs under the signed-in user's session.
- Native SharePoint Person fields are selected/expanded into stable IDs, display names, and email addresses and are written via `FieldNameId` lookup values.
- Issues and progress logs are read with PnPjs page iteration instead of a silent 5,000-row cap.
- `SharePointRoleResolver.ts` maps the current user's SharePoint groups to Admin, Quality Manager, Task Owner, or Reader. Resolution is fail-closed; only localhost gets the explicit Admin development override.
- Connection Diagnostics validates access, schema, choices, indexes, and a required-field-safe create/update/delete round trip with cleanup for both Lists.
- Recharts, Lucide, and generated Tailwind utilities are bundled in the `.sppkg`; production loads no Tailwind CDN.
- SharePoint direct remains the default. Local development in direct mode uses `MockDataService`; tenant direct mode uses `SharePointDataService`. Explicit backend mode uses `BackendApiDataService` in either environment.
- The [integration contract](../backend/sharepoint/qstar-sharepoint-graph-integration.md#4-service-behavior) defines partial writes and version checks; the [provisioning guide](../backend/sharepoint/provisioning/README.md) owns reference allocation and journal upgrades.
- See the [user guide](../README.md#working-with-issues) for lifecycle actions and [saving and draft recovery](../README.md#saving-and-recovering-drafts) for conflicts, busy controls, accepted-write warnings, and retained recovery copies.
- Date-only values use calendar arithmetic and local formatting. SharePoint date envelopes are normalized for native date inputs and reminder comparisons.
- Unit, service-contract, and real React interaction regressions live in `tests/`.
- The Settings screen displays the actual connection. Direct mode uses the web-part site/list properties; backend mode uses the authoritative `/me` response. Changing mode, target, resource or identity remounts the app and resolves access again, while unrelated renders reuse the existing services.

The original source and standalone preview remain in `prototype/` as the requirements/reference baseline.
The port in `QstarPrototype.tsx` currently disables ESLint and TypeScript checking with file-level directives; passing those checks does not establish that component's lint or type safety. Its interaction coverage is in `tests/ui-regressions.test.ts`.

## Requirements

- Node.js `>=18.17.1 <19.0.0`
- A trusted SPFx development certificate for `gulp serve`
- For tenant testing: a Q-Star development site provisioned for the selected [beta or production access model](../backend/sharepoint/provisioning/README.md)

## Install, test, and build

Run these commands from `frontend/` with the Node version required above:

```bash
npm ci
npm test
npx gulp bundle --ship
npx gulp package-solution --ship
```

`npm test` regenerates the locally bundled utility stylesheet, runs the unit suite, then runs SPFx lint, TypeScript, Sass, and webpack checks.

The deployable package is:

```text
sharepoint/solution/qstar-issue-manager.sppkg
```

## Local development

```bash
npx gulp trust-dev-cert   # one-time; macOS may request an administrator password
npx gulp serve --nobrowser
```

Then open the tenant SharePoint Workbench with the debug-manifest query printed by `gulp serve`. Localhost uses mock data and a clearly marked development Admin role.

## Tenant work remaining

1. Provision a dedicated development site with `backend/sharepoint/provisioning/`.
2. For beta, use the explicit beta provisioning entry point and enable **Beta access mode** in the web part properties. Existing site Owners map to Admin, Members/editors to Quality Manager, and read-only visitors to Reader.
3. Run Connection Diagnostics in the real tenant.
4. Build the assignment-permission, intake, and reminder flows from `backend/power-automate/qstar-power-automate-flows.md`.
5. Validate whether nested Entra groups are enumerated through SharePoint; add the documented `MSGraphClientV3` fallback only if required.
6. Run the role/permission/UAT checklist before App Catalog production deployment.

Choose the entry point and review existing-data migrations in [the provisioning guide](../backend/sharepoint/provisioning/README.md).

## Optional backend mode

**Data source** defaults to `sharepoint`. Selecting `backend` is an explicit opt-in; configure an HTTPS API base URL and the Entra application ID/resource URI in the web-part properties. The server ships with `qstar.backend.enabled=false` and must remain disabled until the backend's local and tenant acceptance checklist is complete. See [the backend deployment guide](../backend/service/README.md) for delegated permissions, consent, exact tenant CORS origin and activation. No live tenant acceptance is implied by frontend tests.

Backend access comes only from `/me`: its verified role, SharePoint user and configured site/list targets are used by the app. Missing identity/configuration or failed access resolution blocks the app without switching to direct SharePoint. Settings failures remain visible. Backend diagnostics use `/diagnostics`; direct diagnostics remain unchanged.

The transport preserves person IDs/emails, calendar dates, reminder cycles and ETags. PATCH sends `If-Match` and accepts either the saved issue or a saved-but-unreadable receipt; 412 enters the existing conflict/reload flow. Creates and progress appends retain server IDs and warnings when readback fails. A known successful response that cannot identify its record consumes the submitted form/note and blocks that operation until **Reload saved data** succeeds. The accepted content must not be submitted again. These accepted-write guards are scoped to the connection and user in browser `sessionStorage`, so they survive same-tab page reloads as well as navigation and component remounts; if browser storage is unavailable, the in-memory guard still survives component remounts. This differs from unsaved draft archives, which remain only in the mounted app session. Recovery does not discard unrelated drafts or archived evidence.

Canonical region choices include **France (Paris)** and **Asia Pacific (Bangkok)**; historical `Asia Pacific` values normalize to Bangkok alongside the existing legacy aliases.
