# Backend — data & automation layer

There is no custom API server for this system. The "backend" is Microsoft 365 itself:

- **`sharepoint/`** — the SharePoint List that stores every issue, the field/column mapping, and the Microsoft Graph read/write data layer design that the frontend calls into (delegated auth, signed-in user, `Sites.Selected` scope only).
  - `qstar-sharepoint-graph-integration.md` — column reference and Graph data layer design.
  - `provisioning/` — separate beta entry points (`provision-qstar-beta.ps1` / `provision-qstar-beta-m365.sh`) for lists only, and production entry points (`provision-qstar.ps1` / `provision-qstar-m365.sh`) that also create role groups and permissions.
- **`power-automate/`** — the intake flow (Microsoft Form → List item) and the daily reminder flow, documented step-by-step for building in the Power Automate designer.

See [`../CLAUDE.md`](../CLAUDE.md) for the full production constraints (no public-internet calls, single-site scope, no secrets in the frontend, data stays in the tenant).
