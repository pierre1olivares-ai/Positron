# Q-Star beta change log

**Branch:** `beta`

**Draft pull request:** [#1 — Prepare Q-Star beta deployment](https://github.com/pierre1olivares-ai/Positron/pull/1)

**Prepared:** 14 July 2026

## Purpose

This document records the recommendations made during the Q-Star review, what was implemented and pushed to the `beta` branch, and what still has to be completed in the Microsoft 365 tenant. It is intended as the technical handoff for colleagues reviewing or deploying the beta.

## Recommended architecture

| Area | Recommendation | Beta implementation | Production direction |
|---|---|---|---|
| Hosting | Host Q-Star as an SPFx web part on one dedicated SharePoint Online site. | Use a blank Communication site. No separate application server is required. | Keep the dedicated-site boundary unless a future customer-hosted API is approved. |
| Package distribution | Limit the beta package to the Q-Star site. | Upload the `.sppkg` to a Site Collection App Catalog on the Q-Star site. | Use the tenant App Catalog only if wider organizational distribution is approved. |
| Authentication | Use the existing Microsoft 365 sign-in rather than building a separate username/password system. | The web part runs as the signed-in SharePoint user. No client secret or separate identity database is used. | Continue with Entra-backed identity and Conditional Access. |
| Authorization | Treat React role checks as presentation; SharePoint permissions are the security boundary. | Beta access mode maps existing site permissions: Owners to Admin, Members/editors to Quality Manager, and read-only visitors to Reader. | Use the four dedicated Q-Star groups and item-level Task Owner permissions. |
| Data | Keep operational data in SharePoint lists on the same site. | PnPjs reads and writes the Q-Star lists using the current user's SharePoint session. | Retain the same-site model unless an approved API becomes necessary. |
| People fields | Use native SharePoint Person columns rather than parallel text/email fields. | Reporter, Task Owner, Verifier, and progress Author use stable SharePoint lookup IDs, display names, and email addresses. | Keep native Person fields in production. |
| Automations | Move background work out of the browser. | Intake, assignment-permission, and reminder/notification flows are documented. | Create and validate all three Power Automate flows before production. |
| Front-end assets | Avoid public runtime dependencies. | Tailwind CSS, Lucide icons, and Recharts are bundled into the SPFx package. | Continue the no-public-CDN policy. |

## Changes implemented on the beta branch

### SharePoint Framework application

- Ported the validated Q-Star prototype into the SPFx web part.
- Replaced prototype storage calls with the `IDataService` abstraction and the real SharePoint data service.
- Added loading, save-error, and access-resolution states.
- Uses mock data and a development Admin role only when served from localhost.
- Uses the current Microsoft 365 user identity in SharePoint Online.
- Accepts real Microsoft 365 names and email addresses for Task Owners and Verifiers.
- Prevents unresolved Person assignments from being saved.

### SharePoint data layer

- Added native Person-field `$select`/`$expand` reads and `FieldNameId` writes.
- Added `ensureUser` handling for Microsoft 365 email addresses.
- Added stable Person IDs and emails to the issue and progress models.
- Added asynchronous paging so list reads are not limited to the first page.
- Corrected blank-date clearing, sorting, required fields, settings errors, and progress-author writes.
- Re-fetches newly created issues so SharePoint-calculated values are returned to the UI.

### Access modes

- Added production role resolution with this priority: Admin, Quality Manager, Task Owner, Reader.
- Added explicit **Beta access mode** to the web-part properties; it is enabled in the beta's default web-part configuration.
- Beta mode uses effective site permissions and does not require Q-Star groups.
- Production mode remains fail-closed to Reader when no Q-Star group matches.

### Separate provisioning profiles

The repository supports two logical profiles, each with a PowerShell and Microsoft 365 CLI entry point.

**Beta — lists and fields only; no Q-Star groups or role assignments:**

- `backend/sharepoint/provisioning/provision-qstar-beta.ps1`
- `backend/sharepoint/provisioning/provision-qstar-beta-m365.sh`

**Production — lists, fields, four Q-Star groups, and site permissions:**

- `backend/sharepoint/provisioning/provision-qstar.ps1`
- `backend/sharepoint/provisioning/provision-qstar-m365.sh`

Only one toolchain should be used for a deployment; do not run both PowerShell and the CLI script.

### Diagnostics and testing

- Repaired Connection Diagnostics so temporary issue and progress records satisfy required fields.
- Diagnostics now use a real SharePoint Person lookup ID and clean up both temporary records.
- Added tests for role priority, beta permission mapping, least-privilege fallback, Person parsing, field mappings, and diagnostic payloads.
- Added local Tailwind generation to the build and test commands.

### Documentation

- Updated the implementation checklist, SharePoint integration design, frontend guide, backend guide, Power Automate flow guide, and repository handoff notes.
- Added the one-page product proposal at `outputs/qstar-one-page-proposal.md`.

## Beta deployment sequence

1. Create a dedicated blank SharePoint Online Communication site.
2. Restrict site membership to the beta participants and disable external sharing.
3. Ask a SharePoint administrator to enable a Site Collection App Catalog on that site.
4. Run one beta provisioning entry point against the site. This creates the three Q-Star lists and required fields without creating groups.
5. Build the SPFx package with Node 18, or use the locally generated package from `frontend/sharepoint/solution/qstar-issue-manager.sppkg`.
6. Upload and enable the package in the Q-Star site's Site Collection App Catalog.
7. Install the app on the site if SharePoint requests it, then add `QstarIssueManager` to a modern page.
8. Leave the web part's SharePoint site URL blank, retain the default list names, and keep Beta access mode enabled.
9. Run Connection Diagnostics, then execute the beta UAT checklist.

The `.sppkg` is a generated build artifact and is not committed to Git. To rebuild it:

```bash
cd frontend
nvm use 18
npm ci
npm run styles:prototype
npx gulp clean
npx gulp bundle --ship
npx gulp package-solution --ship
```

## Validation completed locally

- Six core rule tests pass.
- SPFx lint, TypeScript, Sass, and Webpack test build pass on Node 18.
- Clean `bundle --ship` and `package-solution --ship` complete successfully.
- The Microsoft 365 CLI production and beta entry points pass `bash -n` syntax checks.
- The generated package contains a single current hashed web-part bundle.
- No Tailwind CDN, jsDelivr, or unpkg runtime references were found in the built assets.

## Known beta limitations and remaining work

- The beta does not provide a distinct Task Owner security role. Site Members/editors are treated as Quality Managers and can edit all beta issues.
- Item-level assignment permissions are not enforced until the assignment-permission flow and production group model are deployed.
- The three Power Automate flows are documented but have not been created or tested in the tenant.
- Live SharePoint provisioning, Connection Diagnostics, role checks, paging, and end-to-end UAT still require the beta tenant.
- Nested Entra-group behavior has not been validated; no Graph fallback should be added unless tenant testing proves it is necessary.
- The local browser workbench was not visually exercised because trusting the SPFx development certificate requires a local administrator action.
- The PowerShell provisioning entry points were reviewed but could not be parser-tested locally because `pwsh` is unavailable.
- `QstarPrototype.tsx` is currently a large port under `@ts-nocheck`; splitting and fully typing it is recommended before broad production rollout.

## Commits included before this change log

- `eb6c375` — Prepare Q-Star beta deployment
- `94a8a96` — Add Q-Star one-page proposal

The draft pull request from `beta` to `main` remains the review location for the complete change set.
