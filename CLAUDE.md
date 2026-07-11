# Q‑Star Issue Manager — project context (read me first)

This file primes Claude Code to continue a project that began in a Claude chat. Read it, then see the file inventory and "Next steps" below.

## What this is
A quality‑issue management tool for **time:matters** (logistics; Lufthansa Cargo group), built around **ISO 9001:2015**. It covers the full lifecycle: report → triage → assign → progress → on‑hold → (for non‑conformities) a 2‑month effectiveness test → close → re‑open. It has a dashboard, an issue register, role‑based access, and email reminders.

## Current state
- A **working prototype** exists as a **single‑file React app**: [`src/qstar-issue-manager.jsx`](src/qstar-issue-manager.jsx). It currently persists to an injected `window.storage` shim (browser localStorage in the standalone build).
- A **self‑contained preview** exists: [`docs/qstar-live.html`](docs/qstar-live.html) (React + Recharts + Lucide bundled; Tailwind via a public CDN — see constraint below).
- The prototype is feature‑complete for review and has been validated with the Quality team.

## The goal now: productionize into Microsoft 365 / SharePoint
Turn the prototype into a maintainable, deployed tool:
1. Put the code under version control (this repo).
2. Repackage as an **SPFx web part** (SharePoint Framework) that runs inside a **single dedicated SharePoint site**.
3. Replace `window.storage` with a **SharePoint List** data layer (read/write via SPFx/Graph, acting as the signed‑in user).
4. Resolve the four roles (Admin / Quality Manager / Task Owner / Reader) from **Microsoft 365 (Entra ID)** — ideally security groups.
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
- `src/qstar-issue-manager.jsx` — the full prototype app (the starting code; will become the SPFx web part source).
- `docs/qstar-live.html` — clickable preview for demos/requirements.
- `docs/qstar-sharepoint-graph-integration.md` — SharePoint List column map + the read/write data layer design.
- `scripts/provision-qstar.ps1` / `scripts/provision-qstar-m365.sh` — scripts that create the List columns.
- `docs/qstar-power-automate-flows.md` — step‑by‑step build of the intake + reminder flows.
- `docs/qstar-implementation-checklist.md` — the non‑developer rollout plan and IT ask list (governance, security, sequence).

## Suggested next steps (for Claude Code)
1. ~~Initialise the repo and commit the current files.~~ Done.
2. Scaffold an **SPFx web part** solution (Yeoman/`@microsoft/generator-sharepoint`), Node LTS.
3. Port the React component from `src/qstar-issue-manager.jsx` into the web part.
4. Implement the SharePoint List data layer per `docs/qstar-sharepoint-graph-integration.md` (replace `window.storage`).
5. Remove the Tailwind CDN; bundle styling locally.
6. Wire role resolution to Entra groups.
7. Confirm the Power Automate flows against the provisioned List.
8. Build, test in a dev site, then follow `docs/qstar-implementation-checklist.md` for deployment.

> Ask me (Claude Code) to start with any step — e.g. "scaffold the SPFx solution and port the component," or "build the SharePoint List data layer to replace window.storage."
