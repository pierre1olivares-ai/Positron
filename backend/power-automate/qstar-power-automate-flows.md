# Q-Star — tenant workflow implementation

These are build instructions, not deployed flows. Create the three flows in the tenant with an approved service identity in **Q-Star Admins**. Site Full Control alone does not grant access through unique list permissions. Use real list IDs and RootFolder.ServerRelativeUrl; list titles do not necessarily equal URL segments.

Deploy the updated web part, provisioning and intake flow together during a maintenance window. See [provisioning upgrades](../sharepoint/provisioning/README.md). Production uses native Person columns. Beta retains existing site permissions and cannot validate production owner isolation or append-only history.

## A — Microsoft Forms intake

1. Trigger **When a new response is submitted**, then **Get response details**. Enable trigger concurrency **1** for predictable response retries; numbering does not depend on this setting.
2. Read the single Config item's numeric **ReferenceOffset**. Abort and alert if absent, nonintegral or below 1000. Never recalculate it or read `max(QsNumber)` in the flow.
3. **Create item** in Issues, leaving QsNumber empty. Map:

   | Field | Value |
   |---|---|
   | Title, ShortSummary | Short summary |
   | Description | Description |
   | Severity, DepartmentBU, Region | Canonical choice value |
   | ReportedBy Claims | Respondent Microsoft 365 identity |
   | ReportDate | Current business calendar date |
   | ImmediateAction, AlreadyInContact, DeviationType, Origin, AdditionalComments | Corresponding answers |
   | Triaged, TaskCreated, OwnerUpdate | No |
   | Status | Empty |
   | ReminderCycle | initial |

   Normalize old Region answers using `provisioning/region-schema.json` or update the Form before restarting intake. Germany maps to Western Europe (Amsterdam). Unknown values need an explicit business decision.
4. Compute `add(int(<ReferenceOffset>),int(outputs('Create_item')?['body/ID']))` and update **that item's** QsNumber. Supply any connector-required fields from Create item, never blank them. Existing nonempty references remain unchanged. Retry a failed second action by its recorded created item ID; never create another issue. The web part derives the same reference while materialization is pending.
5. Keep a service-owned response-ID → created-item-ID intake log. A manual replay must resume an existing response rather than create a duplicate. Reconcile ambiguous Create timeouts before replaying them.

SharePoint allocates IDs, so browsers and Forms use the same reference scheme. ReferenceOffset was initialized above all legacy references during maintenance and must never change after first use.

## B — Reminders and accepted update notifications

### Setup and calendar contract

Schedule one dispatcher per site, e.g. 07:00 business time. Enable trigger concurrency **1**, keep Apply to each concurrency **1**, and route manual retries through the same dispatcher. Configure `varQualityTeam` and a Windows business time zone, e.g. `W. Europe Standard Time`. Define varToday:

```text
formatDateTime(convertTimeZone(utcNow(),'UTC',variables('varBusinessTimeZone')),'yyyy-MM-dd')
```

Business dates use their stored `yyyy-MM-dd` component, matching the web part. Event timestamps (OwnerUpdateAt, Created) retain their UTC timestamp. Include boundary-time cases when validating site regional settings.

Create a service-owned **Q-Star Reminder Log**: EventKey (Text, indexed), RecipientKey (Text, indexed), State (Choice Pending/Sent), SentAt (DateTime), IssueId (Number, indexed). Only service/Admin accounts may write. An optional **BU Leads** list maps BU Title to LeadEmail; fall back to the Quality Team when absent.

### Event identity and delivery

Each recipient gets a separate row. RecipientKey is `toLower(trim(<actual email>))`.

| Source | EventKey |
|---|---|
| Scheduled milestone | `issue:<ID>:cycle:<ReminderCycle or initial>:<RuleKey>:<milestone date>` |
| Owner status update | `issue:<ID>:update:<OwnerUpdateAt ISO timestamp>` |
| Accepted progress entry | `progress:<Progress Log item ID>` |

ReminderCycle changes on reopening. A new cycle, due/hold/test-end date, or recipient therefore permits a new delivery. Do not deduplicate using only IssueId/RuleKey or the day the email was sent.

For each `(EventKey, RecipientKey)`:

1. Query Reminder Log by **both** fields. Escape apostrophes in OData literals by doubling them before constructing the filter.
2. Sent → skip. Pending → alert for operator reconciliation; do not automatically resend when the previous delivery outcome is unknown.
3. Absent → create Pending, send Outlook email, then update that log ID to Sent/SentAt. For a known send failure, retry the same Pending delivery only after confirming nothing was delivered.

Email and SharePoint are not one transaction; this exposes the ambiguous send/log window rather than promising exactly-once delivery. Maintain serialized dispatch and do not run competing flows against this log. At rollout set an event cutoff or mark older events handled to avoid emailing the historical journal. Retain dedupe records while their lifecycle events can still recur.

### Update events, independent of active-task reminders

Process these even when the issue has just closed or changed owner:

- Page through accepted Progress Log entries after the rollout cutoff. Exclude folders (FSObjType = 1). Derive the parent from the exact `<ProgressRoot>/issue-<positive ID>` FileDirRef, never ParentItemId. Built-in Author/Created identify author and event time. Send the accepted text to the Quality Team with `progress:<entry ID>`. A comment is one durable row; it does not need a separate OwnerUpdate patch. Reassignment must not discard the previous owner's accepted comment.
- For issues with OwnerUpdateAt after the cutoff, send their latest status-update text using the timestamp EventKey. Do **not** require OwnerUpdate = Yes: that Boolean controls a UI acknowledgement badge which a QM may dismiss before the schedule. Do not clear OwnerUpdateAt after delivery. It is a latest-update snapshot: multiple status changes between runs produce the most recent notification; every accepted progress entry has its own event ID.

Enable pagination on all Get items actions and set thresholds for actual register/log sizes. Do not hard-limit to 5000. Migrate root-level journal entries before deploying the web part; invalid paths generate an operational alert instead of guessed grouping.

### Classify before computing dates

Active tasks have Triaged = Yes, TaskCreated = Yes, and status Created, In Progress, On Hold or Under Testing/Revision. Apply this ordered branching:

1. **Under Testing/Revision**: require NC Minor/NC Major and nonempty ImplementationDate. Only then calculate testEnd with `formatDateTime(addToTime(<ImplementationDate>,2,'Month'),'yyyy-MM-dd')`. Calendar-month addition clamps month-end, matching the app. Do not evaluate DueDate.
2. **On Hold**: require HoldUntil, then use only hold milestones. Never send due/overdue escalation for a held issue. A missing HoldUntil creates a data-quality alert.
3. **Created/In Progress**: test `empty(DueDate) = false` **before** composing daysUntilDue or formatting a due date in an email body. Missing dates create an alert and skip those rules, not a null formatDateTime call.

For a validated calendar date d, days until it are:

```text
div(sub(ticks(concat(d,'T00:00:00Z')),ticks(concat(variables('varToday'),'T00:00:00Z'))),864000000000)
```

### Milestones

Use the milestone date in EventKey, not varToday. Deliver when due or past and not yet logged, so missed runs can catch up.

| Branch / RuleKey | Milestone | Recipient |
|---|---|---|
| Active / OWNER_HEADSUP | DueDate − 3 days | Owner |
| Active / OWNER_DUE | DueDate | Owner |
| Active / OWNER_OVERDUE_3 … OWNER_OVERDUE_30 | DueDate + 3,6,…,30 days | Owner |
| Active / QM_ALERT | DueDate + 1 day | Quality Team |
| Active / BU_ESCALATION | DueDate + 7 days | BU lead or Quality Team |
| Hold / HOLD_SOON | HoldUntil − 7 days | Owner |
| Hold / HOLD_END_OWNER | HoldUntil | Owner |
| Hold / HOLD_END_QM | HoldUntil | Quality Team |
| NC / NC_TEST_SOON | testEnd − 7 days | Owner |
| NC / NC_TEST_DONE_OWNER | testEnd | Owner |
| NC / NC_TEST_DONE_QM | testEnd | Quality Team |

Resolve owner email from the native Person field. Missing owner email creates a Quality Team alert and must not be marked delivered. For long outages, optionally group catch-up milestones into one email per recipient and mark only its included EventKeys sent. HTML-encode issue text in email bodies. Use the configured tenant page and issue ID for links; never include credentials.

## C — Reconcile assignment and journal permissions

Production requires this flow and the provisioned root ACLs. **Q-Star Append Progress** grants read + Add Items without Edit Items/Delete Items/Manage Permissions. React is presentation: SharePoint Edit does not provide column-level authorization.

1. Trigger **When an item is created or modified** on Issues; set concurrency **1**. Do not require a populated TaskOwner or skip using editable PermissionedOwnerEmail. Initial unassigned issues and clearing ownership must run. No permission marker write-back is needed.
2. Fetch the **latest** issue after the run starts, rather than using stale trigger ownership. Resolve the four group IDs, custom append role ID, executing service principal ID, and actual Progress RootFolder.ServerRelativeUrl.
3. Ensure `<ProgressRoot>/issue-<ID>` exists. Provisioning enables folders/content types. On a concurrent-create conflict, fetch that exact folder before continuing. Never use a caller-provided path/ParentItemId.
4. Reconcile both the issue item and folder ListItemAllFields. Read HasUniqueRoleAssignments. Only when false POST `breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)`. Then enumerate RoleAssignments with Member IDs/RoleDefinitionBindings.
5. Add missing desired bindings, then explicitly remove **every binding not in the matrix**, including former owners and broad Members grants. Retain SharePoint-managed Limited Access (1073741825) on ancestors; it does not grant item read/edit rights itself:

   ```text
   POST <object>/roleassignments/addroleassignment(principalid=<id>,roledefid=<role>)
   POST <object>/roleassignments/removeroleassignment(principalid=<id>,roledefid=<role>)
   ```

   Objects are `/_api/web/lists(guid'<issues-id>')/items(<ID>)` and `/_api/web/GetFolderByServerRelativePath(decodedUrl='<folder>')/ListItemAllFields`. Escape URLs/OData literals and use retrieved paths.

   | Principal | Issue | Journal folder |
   |---|---|---|
   | Admins and executing service identity | Full Control (1073741829) | Full Control |
   | Quality Managers | Edit (1073741830) | Q-Star Append Progress |
   | Task Owners group | Read (1073741826) | Read |
   | Readers | Read | Read |
   | Current TaskOwner, if populated | Edit | Q-Star Append Progress |

   If one principal appears twice, retain its intended role union. Never remove Admin/QM authorization because that person used to own the issue.
6. Empty TaskOwner means no individual owner grant on either object. Repeating breakroleinheritance on an already-unique object does **not** clear its ACL; explicit removal is essential.
7. Confirm both ACLs, alert on failure, and retry full reconciliation using the latest issue. UI assignment and asynchronous permissions are not atomic: former access may persist until this flow completes. Validate required revocation latency before rollout.

Entries inherit folder permissions. Owners append only to their assigned folder and read elsewhere; they cannot edit/delete earlier entries. The service uses server-managed FileDirRef and Author/Created, not writable ParentItemId. Migrate loose legacy rows in place before deployment, preserving identity/history.

## Required tenant checks

- Concurrent Forms/browser intake produces distinct references. Retrying materialization preserves one item and one reference.
- Owner A can append in A's folder; REST append to B's folder and edit/delete existing entries fail. Forged ParentItemId cannot redirect grouping.
- A→B and A→empty remove A's issue Edit/folder Append after reconciliation; Admin/QM retain access. Repeat the flow and queue rapid reassignments.
- On Hold with old DueDate sends only hold milestones. NC under test with blank DueDate sends NC milestones. Missing required branch dates create alerts without failing expressions.
- A comment creates one row and one QM event. Rerunning notifications does not duplicate it. Exercise email/log failure and Pending reconciliation.
- Reopen without changing DueDate; change owner/due date/hold date/test date: new cycle/recipient/milestone messages can send, while repeated identical events are suppressed.

References: [permission inheritance](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/95bd3a91-79d4-418a-9d4a-f6aac0db44c5), [creating items in folders](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest#create-list-item-in-a-folder), [flow concurrency](https://learn.microsoft.com/en-us/power-automate/guidance/coding-guidelines/optimize-power-automate-triggers), [date functions](https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference).
