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
- Local development uses `MockDataService`; tenant builds use `SharePointDataService`.
- Saves send edited fields with the version read when editing began. Conflicts preserve the draft and offer an explicit reload; an accepted write followed by a failed refresh is not treated as a retryable submission.
- New QS references use an immutable configured offset plus the SharePoint item ID. Existing references remain unchanged; browser and Forms intake use the same allocation rule.
- Progress entries are appended under secured `issue-{id}` folders. SharePoint supplies author/time; folder paths determine the parent issue.
- Every status save passes lifecycle validation, including the NC test period and verifier requirements. OFIs can close without NC-only fields.
- Date-only values use calendar arithmetic and local formatting. SharePoint date envelopes are normalized for native date inputs and reminder comparisons.
- Unit, service-contract, and real React interaction regressions live in `tests/`.
- The Settings screen displays the actual connection. Change its site/lists in the web-part properties; changing the target remounts the app and resolves the role again.

The original source and standalone preview remain in `prototype/` as the requirements/reference baseline.

## Requirements

- Node.js `>=18.17.1 <19.0.0`
- A trusted SPFx development certificate for `gulp serve`
- For tenant testing: a provisioned Q-Star development site and the four Q-Star SharePoint groups

## Install, test, and build

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

Use the explicitly named beta provisioning entry point for the pilot:

- PowerShell: `provision-qstar-beta.ps1`
- Microsoft 365 CLI: `provision-qstar-beta-m365.sh`

The production entry points (`provision-qstar.ps1` and `provision-qstar-m365.sh`) create the lists plus all four Q-Star role groups and reconcile list, issue, and progress-folder permissions. Review existing-data migrations first in [the provisioning guide](../backend/sharepoint/provisioning/README.md).
