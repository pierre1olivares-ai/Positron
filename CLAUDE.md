# Q‑Star Issue Manager — project context (read me first)

This file primes Claude Code to continue a project that began in a Claude chat. Read it, then see the file inventory and "Next steps" below.

## What this is
A quality‑issue management tool for **time:matters** (logistics; Lufthansa Cargo group), built around **ISO 9001:2015**. It covers the full lifecycle: report → triage → assign → progress → on‑hold → (for non‑conformities) a 2‑month effectiveness test → close → re‑open. It has a dashboard, an issue register, role‑based access, and email reminders.

## Current state
- A **working prototype** exists as a **single‑file React app**: [`frontend/prototype/qstar-issue-manager.jsx`](frontend/prototype/qstar-issue-manager.jsx). It currently persists to an injected `window.storage` shim (browser localStorage in the standalone build).
- A **self‑contained preview** exists: [`frontend/prototype/qstar-live.html`](frontend/prototype/qstar-live.html) (React + Recharts + Lucide bundled; Tailwind via a public CDN — see constraint below).
- The prototype is feature‑complete for review and has been validated with the Quality team.
- The repo is split into `frontend/` (SPFx web part) and `backend/` (SharePoint List + Power Automate — see `backend/README.md`).
- The **production SPFx web part is implemented and packages cleanly** (`frontend/`, SPFx 1.20, React 17, Node 18). The validated UI is ported, Tailwind/Recharts/Lucide assets are bundled locally, data runs through SharePoint REST/PnPjs with native Person fields and paging, SharePoint role resolution is wired, diagnostics are repaired, and an initial automated unit suite exists. Remaining work requires a real tenant: provisioning, nested-group validation, Power Automate creation, permission/UAT checks, and App Catalog deployment.

## Changes made

### 2026-07-14 — Authentication and authorization architecture documented

This section records the approved design decision. Implementation status is recorded immediately below it.

- Do **not** build a separate Q-Star login, password store, token store, or browser session system. SPFx runs inside SharePoint and uses the user's existing Microsoft Entra ID sign-in, MFA, Conditional Access, and account lifecycle.
- Keep authentication and authorization separate: Entra ID/SharePoint establishes who the user is; Q-Star roles and SharePoint permissions decide what that user may do.
- Create four site-contained SharePoint groups: **Q-Star Admins**, **Q-Star Quality Managers**, **Q-Star Task Owners**, and **Q-Star Readers**. Where IT governance requires centrally managed membership, place the corresponding Entra security groups inside the SharePoint groups.
- Resolve the UI role from SharePoint group membership/effective site permissions. Validate nested Entra-group resolution in the real tenant; if SharePoint does not enumerate nested membership reliably, use SPFx's authenticated `MSGraphClientV3` only for the group-membership lookup, with the smallest approved delegated permission. Do not introduce a second login or a client secret.
- Treat React role checks as presentation only. SharePoint list/item permissions are the security boundary. Admins and Quality Managers can edit all issues, Readers are read-only, and an assigned Task Owner receives edit permission only on their assigned issue. Power Automate should update item permissions when ownership changes.
- Continue using SharePoint REST/PnPjs with `SPFx(context)` for same-site data. This uses the signed-in user's SharePoint session and needs no app registration or Graph site permission.
- Standardize people fields as native SharePoint Person columns. Read expanded identity values (`Id`, display name, email) and write lookup IDs such as `TaskOwnerId`; do not mix that mode with optional `*Email` text companion columns.
- Remove production dependence on the prototype's email-to-role table and its tenant/client-ID settings. They may remain only in an explicitly marked local demo build.

### 2026-07-14 — Local production implementation completed

- Ported the validated 1,800+ line UI into SPFx and replaced `window.storage` with `IDataService` operations.
- Added fail-closed SharePoint-group role resolution; the simulated identity picker and production email-to-role table are removed.
- Standardized native Person fields, lookup-ID writes, stable identity values, blank-date clearing, settings error reporting, and paged List reads.
- Repaired diagnostics to populate required fields and clean up both issue and progress items.
- Aligned Bash and PowerShell provisioning, including role groups, permission levels, indexes, business-unit choices, and the permission-flow marker field.
- Bundled Recharts, Lucide, and Tailwind-generated styles locally with no Tailwind CDN runtime.
- Added unit tests for role priority, least-privilege fallback, Person values, field maps, and diagnostic payloads.
- Added the item-assignment permission flow design and produced a clean production `.sppkg`.
- Tenant provisioning, live diagnostics, flow creation, nested Entra-group validation, UAT, and deployment remain external tasks.
- Added explicitly named beta provisioning entry points that create lists only; production entry points create the four role groups. Beta access mode maps existing site permissions to Admin, Quality Manager, and Reader.

## The goal now: productionize into Microsoft 365 / SharePoint
Turn the prototype into a maintainable, deployed tool:
1. Put the code under version control (this repo).
2. Repackage as an **SPFx web part** (SharePoint Framework) that runs inside a **single dedicated SharePoint site**.
3. Replace `window.storage` with a **SharePoint List** data layer (read/write via SPFx and SharePoint REST/PnPjs, acting as the signed‑in user).
4. Resolve the four roles (Admin / Quality Manager / Task Owner / Reader) through **SharePoint site groups**, backed by Entra security groups where required, and enforce access with SharePoint permissions.
5. Move reminders/notifications to **Power Automate** (intake flow + daily reminder flow); intake via a **Microsoft Form**.

### Hard constraints for production
- **No public‑internet calls.** Bundle all assets locally (the preview's Tailwind CDN must be removed). The app must load nothing from external sites.
- **Contained to one site.** If Graph is used, request **`Sites.Selected`** scoped to the Q‑Star site only — never tenant‑wide scopes.
- **No secrets in the front‑end.** Use the user's delegated identity.
- **Data stays in the tenant** (SharePoint List).

## Domain notes Claude Code should know
- **Roles & tabs:** Admin (full + IT settings), Quality Manager (full, no IT), Task Owner (My tasks first; Dashboard, Register, Reminders), Reader (Dashboard only).
- **Statuses:** Created, In Progress, Under Testing/Revision (NC‑only 2‑month effectiveness test), On Hold, Closed, Rejected.
- **Categories:** NC (Minor/Major), OFI, Other.
- **Key behaviours:** on‑hold requires reason + resume date and reminds a week before / on the day; owner comments and status changes notify the QM; closed issues are read‑only with a Re‑Open (Admin/QM) that restarts the cycle; the dashboard has a global Department/Region filter and per‑chart date footers.
- Brand: navy `#1B205C` ("Space Cadet") + yellow `#FBB900` (Selective Yellow).

## Repo structure
- `frontend/prototype/qstar-issue-manager.jsx` — the full prototype app (the starting code; will become the SPFx web part source).
- `frontend/prototype/qstar-live.html` — clickable preview for demos/requirements.
- `backend/sharepoint/qstar-sharepoint-graph-integration.md` — SharePoint List column map + the read/write data layer design.
- `backend/sharepoint/provisioning/provision-qstar-beta.ps1` / `provision-qstar-beta-m365.sh` — beta entry points that create lists only.
- `backend/sharepoint/provisioning/provision-qstar.ps1` / `provision-qstar-m365.sh` — production entry points that create lists plus the four role groups and site permissions.
- `backend/power-automate/qstar-power-automate-flows.md` — step‑by‑step build of the intake + reminder flows.
- `docs/qstar-implementation-checklist.md` — the non‑developer rollout plan and IT ask list (governance, security, sequence).

## Recommended build steps

Implementation status, in the required order:

1. **Create the development site — tenant pending.** Run a `provision-qstar-beta*` entry point for lists only. Run a production `provision-qstar*` entry point when the four Q-Star groups and permissions are required.
2. **Reconcile the List schema and data layer — complete locally.** Native Person columns, `$select`/`$expand`, lookup IDs, paging, and script parity are implemented.
3. **Repair and test connection diagnostics — code complete; tenant run pending.** Required payloads and cleanup are fixed and unit-tested.
4. **Implement a role resolver — complete.** Beta access can use effective site permissions; production uses the four Q-Star groups. Nested Entra-group validation and any necessary Graph fallback remain tenant-dependent.
5. **Enforce SharePoint permissions — design/script complete; flow deployment pending.** Site role permissions are provisioned and the assignment-permission flow is documented.
6. **Port the validated React UI — complete locally.** Production uses `IDataService`; localhost alone uses mock data and a development override.
7. **Bundle all UI assets locally — complete.** Tailwind utility CSS, Recharts, and Lucide are packaged with the solution.
8. **Add automated tests — initial suite complete.** Role, identity, mapping, and diagnostics rules are covered; tenant integration and broader status-transition coverage remain.
9. **Build Power Automate workflows — build guide complete; tenant creation pending.** Assignment, intake, and reminders must be created with tenant connections.
10. **Run tenant integration and UAT — pending.** Exercise all roles, reassignment, direct-list permissions, diagnostics cleanup, paging, failure states, and deployment.

> Ask me (Claude Code) to start with any step — e.g. "reconcile the SharePoint Person fields and diagnostics," or "implement the SharePoint role resolver."
