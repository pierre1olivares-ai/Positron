# Backend — data & automation layer

There is no custom API server for this system. The "backend" is Microsoft 365 itself:

- **`sharepoint/`** — the SharePoint List that stores every issue, the field/column mapping, and the Microsoft Graph read/write data layer design that the frontend calls into (delegated auth, signed-in user, `Sites.Selected` scope only).
  - `qstar-sharepoint-graph-integration.md` — column reference and Graph data layer design.
  - `provisioning/` — scripts that create the List and its columns (`provision-qstar.ps1` for Windows/PnP PowerShell, `provision-qstar-m365.sh` for the M365 CLI).
- **`power-automate/`** — the intake flow (Microsoft Form → List item) and the daily reminder flow, documented step-by-step for building in the Power Automate designer.

See [`../CLAUDE.md`](../CLAUDE.md) for the full production constraints (no public-internet calls, single-site scope, no secrets in the frontend, data stays in the tenant).
