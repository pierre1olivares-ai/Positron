# Backend — Microsoft 365 data and automation

SharePoint lists remain the system of record, with tenant Power Automate flows for automation. The SPFx frontend uses same-site SharePoint REST through PnPjs by default. An optional Java 21 gateway uses delegated on-behalf-of access to the same lists and enforces API roles, issue ownership, conditional writes, and lifecycle rules. Both paths retain the caller's SharePoint permissions and native audit identity.

The gateway is disabled by default. Its [service guide](service/README.md) owns authentication configuration, local builds, and the unresolved tenant permission and activation checks. Do not enable it or broaden consent based on local test results alone.

- [Integration contract](sharepoint/qstar-sharepoint-graph-integration.md): fields, native Person identities, reference allocation, secured progress folders, and settings.
- [Provisioning and upgrades](sharepoint/provisioning/README.md): Bash/CLI and PnP PowerShell entry points, reviewed Region/history migration, and beta versus production permissions.
- [Workflow build guide](power-automate/qstar-power-automate-flows.md): Forms intake, scheduled reminders and accepted updates, and assignment/folder ACL reconciliation. These instructions must be implemented and validated in the tenant; no deployed flows are included.
- [Tenant verification](sharepoint/connection-test-plan.md): behavior and permissions that local automated checks cannot establish.

See [production constraints](../CLAUDE.md) for the single-site deployment and tenant data boundary. No application secrets belong in frontend settings.
