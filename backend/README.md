# Backend — Microsoft 365 data and automation

The backend consists of SharePoint lists and tenant Power Automate flows. The SPFx frontend accesses same-site SharePoint REST through PnPjs using the signed-in user's permissions.

- [Integration contract](sharepoint/qstar-sharepoint-graph-integration.md): fields, native Person identities, reference allocation, secured progress folders, and settings.
- [Provisioning and upgrades](sharepoint/provisioning/README.md): Bash/CLI and PnP PowerShell entry points, reviewed Region/history migration, and beta versus production permissions.
- [Workflow build guide](power-automate/qstar-power-automate-flows.md): Forms intake, scheduled reminders and accepted updates, and assignment/folder ACL reconciliation. These instructions must be implemented and validated in the tenant; no deployed flows are included.
- [Tenant verification](sharepoint/connection-test-plan.md): behavior and permissions that local automated checks cannot establish.

See [production constraints](../CLAUDE.md) for the single-site deployment and tenant data boundary. No application secrets belong in frontend settings.
