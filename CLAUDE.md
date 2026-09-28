# Q‑Star Issue Manager — project context (read me first)

This file records project constraints. Use the linked implementation and rollout guides for current status and operating instructions.

## What this is
A quality‑issue management tool for **time:matters** (logistics; Lufthansa Cargo group), built around **ISO 9001:2015**. It covers the full lifecycle: report → triage → assign → progress → on‑hold → (for non‑conformities) a 2‑month effectiveness test → close → re‑open. It has a dashboard, an issue register, role‑based access, and email reminders.

## Current state
- A **working prototype** exists as a **single‑file React app**: [`frontend/prototype/qstar-issue-manager.jsx`](frontend/prototype/qstar-issue-manager.jsx). It currently persists to an injected `window.storage` shim (browser localStorage in the standalone build).
- A **self‑contained preview** exists: [`frontend/prototype/qstar-live.html`](frontend/prototype/qstar-live.html) (React + Recharts + Lucide bundled; Tailwind via a public CDN — see constraint below).
- The prototype is feature‑complete for review and has been validated with the Quality team.
- The repo is split into `frontend/` (SPFx web part) and `backend/` (SharePoint lists, Power Automate, and an optional Java gateway — see `backend/README.md`).
- The running implementation and supported development toolchain are documented in [frontend/README.md](frontend/README.md). Current repair rollout and validation limits are in [the repair handoff](docs/qstar-review-repairs.md); tenant deployment and acceptance remain pending.

## Changes made

### 2026-07-14 — Authentication and authorization architecture documented

This section records the approved design decision. Implementation status is recorded immediately below it.

- Do **not** build a separate Q-Star login, password store, token store, or browser session system. SPFx runs inside SharePoint and uses the user's existing Microsoft Entra ID sign-in, MFA, Conditional Access, and account lifecycle.
- Keep authentication and authorization separate: Entra ID/SharePoint establishes who the user is; Q-Star roles and SharePoint permissions decide what that user may do.
- Create four site-contained SharePoint groups: **Q-Star Admins**, **Q-Star Quality Managers**, **Q-Star Task Owners**, and **Q-Star Readers**. Where IT governance requires centrally managed membership, place the corresponding Entra security groups inside the SharePoint groups.
- Resolve the UI role from SharePoint group membership/effective site permissions. Validate nested Entra-group resolution in the real tenant; if SharePoint does not enumerate nested membership reliably, use SPFx's authenticated `MSGraphClientV3` only for the group-membership lookup, with the smallest approved delegated permission. Do not introduce a second login or a client secret.
- Treat React role checks as presentation only. SharePoint permissions are the security boundary; the [production permission matrix and migration guide](backend/sharepoint/provisioning/README.md#journal-migration-and-production-permissions) owns list, issue, and journal access. [Flow C](backend/power-automate/qstar-power-automate-flows.md#c--reconcile-assignment-and-journal-permissions) owns assignment reconciliation.
- Continue using SharePoint REST/PnPjs with `SPFx(context)` for same-site data. This uses the signed-in user's SharePoint session and needs no app registration or Graph site permission.
- Standardize people fields as native SharePoint Person columns. Read expanded identity values (`Id`, display name, email) and write lookup IDs such as `TaskOwnerId`; do not mix that mode with optional `*Email` text companion columns.
- Remove production dependence on the prototype's email-to-role table and its tenant/client-ID settings. They may remain only in an explicitly marked local demo build.

### Optional backend mode

The Java 21 service under `backend/service/` uses a delegated on-behalf-of token for SharePoint REST. Its API validates the caller's tenant, issuer, audience, scope, and explicit app roles, then applies ownership and lifecycle checks. Backend-mode presentation uses the service's `/me` identity, role, and configured connection. Direct mode retains the SharePoint group design above. The backend is disabled by default; the [service guide](backend/service/README.md) owns activation requirements. Do not substitute app-only access when delegated access fails.

### Implementation handoff

See [frontend/README.md](frontend/README.md) for the implemented app, [the integration contract](backend/sharepoint/qstar-sharepoint-graph-integration.md) for data semantics, and [the repair handoff](docs/qstar-review-repairs.md) for deployment order and evidence limits. The July handoff is retained as a dated [change log](docs/qstar-beta-change-log.md).

## Deployment goal

Complete the [implementation checklist](docs/qstar-implementation-checklist.md) and [tenant verification plan](backend/sharepoint/connection-test-plan.md) before production use. The SPFx port already exists; keep `frontend/prototype/` as historical reference material.

### Hard constraints for production
- **No public‑internet calls.** Bundle all assets locally (the preview's Tailwind CDN must be removed). The app must load nothing from external sites.
- **Contained to one site.** If Graph is used, request **`Sites.Selected`** scoped to the Q‑Star site only — never tenant‑wide scopes. Backend REST activation also requires a verified narrowly scoped delegated permission. A pinned site URL does not itself prove that the token is limited to that site. Keep the backend disabled until IT verifies this boundary; broader consent requires an explicit architecture exception.
- **No secrets in the front‑end.** Use the user's delegated identity. Backend credentials stay server-side and are used only to exchange the authenticated caller's token; no app-only fallback.
- **Data stays in the tenant** (SharePoint remains the system of record; Q-Star requires no separate business database).

## Domain notes Claude Code should know
- **Roles & tabs:** Admin (full + IT settings), Quality Manager (full, no IT), Task Owner (My tasks first; Dashboard, Register, Reminders), Reader (Dashboard only).
- **Statuses:** Created, In Progress, Under Testing/Revision (NC‑only 2‑month effectiveness test), On Hold, Closed, Rejected.
- **Categories:** NC (Minor/Major), OFI, Other.
- **Key behaviours:** on‑hold requires reason + resume date and reminds a week before / on the day; owner comments and status changes notify the QM; closed issues are read‑only with a Re‑Open (Admin/QM) that restarts the cycle; the dashboard has a global Department/Region filter and per‑chart date footers.
- Brand: navy `#1B205C` ("Space Cadet") + yellow `#FBB900` (Selective Yellow).

## Repo structure
- `frontend/prototype/qstar-issue-manager.jsx` — the historical UI reference; the running implementation is in `frontend/src/`.
- `frontend/prototype/qstar-live.html` — clickable preview for demos/requirements.
- `backend/service/` — the optional Java gateway, API contract, security configuration, and local tests.
- `backend/sharepoint/qstar-sharepoint-graph-integration.md` — SharePoint List column map + the read/write data layer design.
- `backend/sharepoint/provisioning/README.md` — beta/production entry points, schema upgrades, reviewed data migrations, and permissions.
- `backend/power-automate/qstar-power-automate-flows.md` — step‑by‑step build of the intake + reminder flows.
- `docs/qstar-implementation-checklist.md` — the non‑developer rollout plan and IT ask list (governance, security, sequence).

## Build and rollout

Use the [frontend commands](frontend/README.md#install-test-and-build) and the [repair deployment order](docs/qstar-review-repairs.md#deployment-order). Track tenant work in the implementation checklist instead of duplicating its status here.
