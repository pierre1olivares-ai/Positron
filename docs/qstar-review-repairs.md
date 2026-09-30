# Review repair handoff

These repairs target the `beta` branch after the initial `main` merge (`04931bb`) and the integration of upstream `70ebc76`. The historical files under `frontend/prototype/` remain reference material. The application runs from `frontend/src/`.

## Changes

- Provisioning entry points surface errors and reconcile existing fields, indexes, current Region choices, and reference configuration. Germany maps to Western Europe (Amsterdam), and Asia Pacific maps to Asia Pacific (Bangkok), only during an explicit Region migration. France (Paris) is available as a new choice.
- See [saving and recovering drafts](../README.md#saving-and-recovering-drafts) for the repaired conflict, busy-state, accepted-write, and recovery behavior.
- SharePoint IDs allocate new QS references through an immutable `ReferenceOffset`. Existing QS references are preserved. Both intake paths use this rule.
- See [working with issues](../README.md#working-with-issues) for the repaired lifecycle actions.
- Production journal entries live in per-issue folders. Current owners and Quality Managers can add/read entries, while SharePoint supplies author/time and folder location determines the issue. Assignment changes remove prior owner grants, including when the owner is cleared.
- Calendar dates retain their day across time zones and DST. Today's reports count in YTD metrics; ISO date envelopes work in date fields and reminder comparisons.
- Automation guidance covers on-hold reminders, accepted owner updates, null dates, and deduplication across dates, recipients, and reopened cycles.
- Settings shows the active web-part connection. Local demo seeds persist through the mock service, and target changes refresh service state and role resolution.
- The optional Java backend uses delegated SharePoint access, strict token and role checks, current-owner restrictions, original ETags, native Person fields and journal folders, stable references, and the same lifecycle rules. Its service guide owns configuration and activation requirements. Direct SharePoint remains the default, and backend activation is blocked pending the required permission verification.

## Deployment order

1. Back up list data and permissions, then schedule maintenance for existing-data migration. Pause intake and permission/reminder flows while reconciling schema and ACLs.
2. Use the [provisioning guide](../backend/sharepoint/provisioning/README.md) to preview Region and legacy journal changes. Resolve orphan records explicitly. Apply the reviewed migrations and provision the selected beta or production access model. Never reinitialize an established `ReferenceOffset`.
3. Update Forms intake, permission reconciliation, reminder rules, and accepted-progress notifications using the [flow guide](../backend/power-automate/qstar-power-automate-flows.md). Use the same site/list targets as the web part.
4. Build and deploy the new package to the test site. Set the connection and access mode in the web-part properties, then run diagnostics as an Admin.
5. Complete the tenant checks below before resuming automation or production intake. Beta site Members retain the pilot's broad edit rights; production folder ACL enforcement requires the production provisioner and assignment flow.
6. For a later backend rollout, follow the [service guide](../backend/service/README.md) and [backend acceptance cases](../backend/sharepoint/connection-test-plan.md#optional-backend-acceptance). Keep it disabled until the narrow permission boundary, delegated audit identity, API roles, and real SPFx integration are verified. Do not use broader consent or app-only access as an implicit workaround.

## Tenant acceptance checks

Use the [tenant verification plan](../backend/sharepoint/connection-test-plan.md), including its editing/lifecycle/recovery scenarios. That plan owns the acceptance cases; local interaction and service-contract tests do not establish tenant results.

## Evidence and limits

The final local checks on 2026-09-30 passed 267 frontend tests, 57 Java tests, and ten
provisioning tests, with no failures or skips. SPFx lint/type/Sass/webpack checks and the shipping
bundle/package, OpenAPI generation, Java 21 compilation, Spotless, and the executable JAR
build also passed. The legacy prototype component retains its file-level static
check exemptions; its behavior is covered by interaction tests. The Java HTTP tests
use signed synthetic tokens and mocked SharePoint/OBO endpoints; they do not
establish tenant consent or access behavior.

The last automated no-mistakes run (`01M3QBET8H805RC1AAX7896HYE`) committed its
two recovery fixes as `26f1f854`, then stopped during re-review because its Codex
CLI reached a usage limit. Its outcome is failed, not a completed gate pass.
The preserved commit was recovered without dropping earlier commits. Local
source review verified both fixes and their regressions, and the complete test,
formatting and build checks above ran afterward. A one-line Java test-format
correction was required for Spotless. No remote CI or deployment result is claimed.

Local checks execute domain rules, React interactions, mocked SharePoint service contracts, and stateful provisioning command doubles, plus the SPFx lint/type/build/package tools. They do not authenticate against Microsoft 365. Power Automate flows are deployment instructions in this repository, not installed flow exports. Tenant ACL behavior, native PowerShell execution, flow delivery/retries, and real SharePoint round trips require the acceptance checks above.

## Retained dashboard history and navigation drafts

Triage, QM detail and editable progress forms now use App-owned drafts, as owner detail already does. Navigating back and reopening an issue during a pending request retains the visible draft and its original clean baseline/ETag. Rejected requests leave that draft available for an explicit retry. Accepted requests consume only unchanged dispatched fields/text; newer edits, omitted fields and immutable recovery copies remain separate. There is no uncontrolled triage/owner/QM/progress editor mode in production. Drafts and recovery copies remain in memory for this App session.

Explicit triage reload refreshes the clean baseline and ETag while retaining unsent fields. Still-untriaged items retain the default classification and calculated due date. Failed reloads leave the draft intact, and recovering another issue's receipt does not replace its original ETag.

The backend now checks HTTP status before reading a mutation response body. If a known successful create or journal append has an unreadable response stream, it returns an accepted/uncertain receipt that prevents automatic resubmission, without inventing identity or audit details. Failures before a successful status and non-2xx responses remain errors. Successful MERGE operations do not read unused response bodies. HTTP regressions cover broken streams, absent identity headers, single-POST behavior, redirects, permission failures and version conflicts.

The backend permits authorized progress updates and journal notes on Rejected issues, matching the direct client. Closed restrictions, native owner identity, owner field restrictions and journal-only access during effectiveness testing remain enforced.

Historical backlog now reads retained native Issues versions when the dashboard is visible, with at most four item reads in parallel. Direct SharePoint and the optional backend exhaust the version collection, pin continuation URLs to the configured collection, and compare the current item before and after reading. The read-only API is `GET /api/v1/issues/{id}/history`. The response contains `issueId`, `eTag`, `current`, `versions` and `complete`; unsupported or missing native fields are preserved rather than defaulted. History reads do not update issue snapshots or block normal list loading/editing. Revision-keyed caches and stale-response checks protect the chart during concurrent changes.

Coverage uses native item `Created`, `Modified`, `OData__UIVersionString`, and version `VersionId`, `VersionLabel`, `Created`, `IsCurrentVersion`, `Triaged`, `TaskCreated` and `Status`. Only supported major-version labels (`1.0`, `2.0`, etc.) and an internally consistent, contiguous retained suffix establish recorded state. Version IDs are identities, not consecutive ordinals. Retention, missing versions, unsupported shapes/labels, invalid fields/timestamps, denied/incomplete paging and changed snapshots leave affected periods unavailable. Native creation time can prove nonexistence; ReportDate and today's ClosedAt/status cannot supply historical coverage. The mock emulates versions for writes in the current service instance; older persisted mock records without retained versions remain unknown.

Backlog includes the current visible register under the current department/region filters, including currently rejected or untriaged records whose historical eligibility may differ. Monthly gaps and unavailable YTD changes are explicit; lines do not bridge gaps. Other classification filters continue to describe current-record cohorts. Report/closure bars use current record dates and do not count every reopen cycle. This does not reconstruct historical department/region membership, deleted/inaccessible items, unversioned SystemUpdate or administrative overwrite operations. No durable ledger, schema, migration, flow or permission changes were added.

Tenant acceptance must confirm version access, actual top-level choice values and metadata, exhaustive continuation behavior, current-version alignment and retention coverage for both direct and delegated REST modes. Local fixtures exercise the protocol and UI, not tenant success. The accepted live Microsoft 365, native PowerShell, ACL, flow and deployment evidence gap remains; the backend stays disabled and tenant verification remains mandatory before activation.

The read contract follows [PnPjs item version history](https://pnp.github.io/pnpjs/sp/items/#get-version-history), the [PnP native version field mapping](https://github.com/pnp/pnpcore/blob/dev/src/sdk/PnP.Core/Model/SharePoint/Core/Internal/ListItemVersion.cs), and [Microsoft's current list-item REST metadata](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest).
