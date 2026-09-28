# Backend — data & automation layer

SharePoint is still the system of record. As of IT's backend template share (Sept 2026), there's
now also a custom Java service that sits in front of it as a thin authenticated gateway — see
[`service/README.md`](service/README.md#architecture-thin-gateway-in-front-of-sharepoint) for why.

- **`service/`** — the Java/Spring Boot backend, scaffolded from time:matters' internal backend
  template. Validates the SPFx web part's Azure AD bearer token, applies business rules, and
  proxies to SharePoint via Microsoft Graph using its own app identity (`Sites.Selected`).
- **`sharepoint/`** — the SharePoint List that stores every issue, the field/column mapping, and
  the Graph data layer design `service/` implements server-side.
  - `qstar-sharepoint-graph-integration.md` — column reference and Graph data layer design.
  - `provisioning/` — scripts that create the List and its columns (`provision-qstar.ps1` for Windows/PnP PowerShell, `provision-qstar-m365.sh` for the M365 CLI).
  - `connection-test-plan.md` — manual checklist for verifying the SharePoint/Power Automate side in a real tenant.
- **`power-automate/`** — the intake flow (Microsoft Form → List item) and the daily reminder flow, documented step-by-step for building in the Power Automate designer.

The frontend can talk to either layer directly — see the `dataSourceMode` web part property in
[`../frontend/README.md`](../frontend/README.md) — until `service/` has somewhere to actually run.

See [`../CLAUDE.md`](../CLAUDE.md) for the full production constraints (no public-internet calls, single-site scope, no secrets in the frontend, data stays in the tenant).
