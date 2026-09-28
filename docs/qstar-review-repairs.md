# Review repair handoff

These repairs target the `beta` branch after merging `main` (merge commit `04931bb`). The historical files under `frontend/prototype/` remain reference material. The application runs from `frontend/src/`.

## Changes

- Provisioning entry points surface errors and reconcile existing fields, indexes, current Region choices, and reference configuration. Germany maps to Western Europe (Amsterdam) only during an explicit Region migration.
- See [saving and recovering drafts](../README.md#saving-and-recovering-drafts) for the repaired conflict, busy-state, accepted-write, and recovery behavior.
- SharePoint IDs allocate new QS references through an immutable `ReferenceOffset`. Existing QS references are preserved. Both intake paths use this rule.
- See [working with issues](../README.md#working-with-issues) for the repaired lifecycle actions.
- Production journal entries live in per-issue folders. Current owners and Quality Managers can add/read entries, while SharePoint supplies author/time and folder location determines the issue. Assignment changes remove prior owner grants, including when the owner is cleared.
- Calendar dates retain their day across time zones and DST. Today's reports count in YTD metrics; ISO date envelopes work in date fields and reminder comparisons.
- Automation guidance covers on-hold reminders, accepted owner updates, null dates, and deduplication across dates, recipients, and reopened cycles.
- Settings shows the active web-part connection. Local demo seeds persist through the mock service, and target changes refresh service state and role resolution.

## Deployment order

1. Back up list data and permissions, then schedule maintenance for existing-data migration. Pause intake and permission/reminder flows while reconciling schema and ACLs.
2. Use the [provisioning guide](../backend/sharepoint/provisioning/README.md) to preview Region and legacy journal changes. Resolve orphan records explicitly. Apply the reviewed migrations and provision the selected beta or production access model. Never reinitialize an established `ReferenceOffset`.
3. Update Forms intake, permission reconciliation, reminder rules, and accepted-progress notifications using the [flow guide](../backend/power-automate/qstar-power-automate-flows.md). Use the same site/list targets as the web part.
4. Build and deploy the new package to the test site. Set the connection and access mode in the web-part properties, then run diagnostics as an Admin.
5. Complete the tenant checks below before resuming automation or production intake. Beta site Members retain the pilot's broad edit rights; production folder ACL enforcement requires the production provisioner and assignment flow.

## Tenant acceptance checks

Use the [tenant verification plan](../backend/sharepoint/connection-test-plan.md), including its editing/lifecycle/recovery scenarios. That plan owns the acceptance cases; local interaction and service-contract tests do not establish tenant results.

## Evidence and limits

Local checks execute domain rules, React interactions, mocked SharePoint service contracts, and stateful provisioning command doubles, plus the SPFx lint/type/build/package tools. They do not authenticate against Microsoft 365. Power Automate flows are deployment instructions in this repository, not installed flow exports. Tenant ACL behavior, native PowerShell execution, flow delivery/retries, and real SharePoint round trips require the acceptance checks above.
