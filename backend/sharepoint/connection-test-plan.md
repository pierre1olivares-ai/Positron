# Tenant verification — SharePoint and Power Automate

Run after schema, data service, or flow changes against a test site with real Admin, Quality Manager, two owner, and Reader accounts. Local tests cannot substitute for SharePoint permissions or actual flow execution.

## Prerequisites

1. Pause writes and flows for an existing register; export a backup. Run Region and progress migration previews, review mappings, then apply as documented in [provisioning](provisioning/README.md).
2. Confirm one Config item, valid immutable ReferenceOffset, canonical native Person fields, current Region choices, folder-enabled Progress Log, and no loose/unmapped history.
3. Deploy the web part and all three [flows](../power-automate/qstar-power-automate-flows.md). Put their service identity in Q-Star Admins. Configure business timezone, quality recipients, rollout cutoff, and serialized delivery.

## Connection diagnostics

Run the Admin connection test. Confirm site/user access, required columns and types, choice values, indexes, and the disposable item create/update/delete round trip. Inspect and remove any test artifact if cleanup failed. Run read paths as QM, Owner, and Reader; Config must load without granting them settings-write permission.

## References and preserved data

- Create issues concurrently from two browsers and Forms. Every new QsNumber equals ReferenceOffset + that item's ID, with no collision with legacy references.
- Interrupt the second number-materialization action. Resume the same created item; no duplicate intake should appear. A Forms replay uses its response-ID log.
- Compare migrated records with the backup: Region aliases map correctly (Germany → Western Europe (Amsterdam)); descriptions, attachments, old QsNumber, progress item IDs, Author and Created remain intact. Repeating provisioning keeps the same offset and journal mapping.

## Editing, lifecycle, and recovery

Verify the [user-facing save and recovery contract](../../README.md#saving-and-recovering-drafts) with real SharePoint sessions. Copy any evidence before refreshing the page, since recovered drafts are session-local.

- Open one issue in two editors. Save in one, then submit a different draft from the other. The stale save must conflict without overwriting saved fields; its draft remains available until explicit reload/discard. Change only an owner's email and confirm the saved native Person identity changes.
- Start an NC effectiveness test, edit a follow-up note, and save. Confirm the accepted test state persists. Exercise both closure routes before and after the test end, including month ends and a verifier identity; close an OFI without NC-only fields. Reopen and verify the new cycle retains the reference and journal.
- While a save, append, or reload is pending, return to the register and reopen that issue. Detail/progress editing and both closure routes must remain blocked. A rejected append keeps its text/error, including with intervening reload requests; closing requires a fresh action after pending work settles. An independent issue remains editable.
- Finish issue A's triage or reload while editing B. B must stay open with its detail and progress drafts intact. A clean editor reopened during an ordinary save must use the accepted version after that save finishes.
- Leave detail, hold, and progress drafts, then reload after another session closes the issue or reassigns it away. Confirm read-only controls and selectable recovery copies. Reopen or reassign back, type and submit fresh content, and repeat the transition: earlier copies, including identical text in distinct copies, must remain available. Discard one copy and verify other copies, live drafts, and saved data remain unchanged. Successful own detail saves and progress posts must clear only their submitted draft part.
- Accept a close through each route, then fail its readback. It must remain closed and read-only across register navigation, with unsubmitted recovery content retained. Trigger a saved warning on B and reload B; return to A and use its own **Reload this issue** action. Exercise a failed reload followed by a successful one without losing archives or allowing writes from the unavailable version.
- Change the connection in web-part properties and confirm Settings displays that target and access is resolved again. Confirm today's reports appear in dashboard year-to-date totals.

## Production permissions and journal integrity

- Create an unassigned issue; Flow C creates its folder. Assign Owner A; A can edit the issue and append inside its exact folder.
- Attempt REST append into Owner B's folder as A. It must fail. A cannot edit/delete a previous entry or alter its server author/time. A forged ParentItemId cannot redirect UI grouping.
- As a QM, append to either issue but verify editing/deleting old journal entries is denied. Admin maintenance access remains intentional.
- Reassign A → B, then B → empty. Confirm each former owner's issue Edit and folder Append are removed after Flow C finishes, while group read and QM/Admin access remain. Repeat reconciliation and rapidly queue changes; final ACLs reflect the latest assignment.
- Confirm a successful comment creates one durable journal item even if a later reload fails. Refresh before manually retrying an ambiguous request. Beta cannot validate production isolation because it retains site permissions.

## Notifications and dates

All scenario items need Triaged = Yes and TaskCreated = Yes. Use explicit business calendar dates, including month ends and a daylight-saving boundary.

| Scenario | Expected |
|---|---|
| In Progress, DueDate ten days ago | Previously unlogged overdue owner, QM and BU milestones catch up |
| On Hold with old DueDate, HoldUntil today | Hold owner/QM events only; no overdue escalation |
| NC Under Testing/Revision, ImplementationDate two calendar months ago, no DueDate | Test-end owner/QM events; no null date failure |
| Missing the branch's required date | Data-quality alert; no malformed email/date expression |
| Accepted progress then reassignment or closure | One Quality Team event for that entry ID |
| Owner status change then QM dismisses badge | Latest OwnerUpdateAt notification still dispatches |
| Reopen with same due date | New ReminderCycle permits new milestones |
| Changed owner or milestone date | New recipient/milestone event can send |

Run identical events twice: the second run skips Sent records. Simulate a send/log interruption: Pending is surfaced for reconciliation and is not blindly resent. Check actual mailbox delivery as well as run history. Verify every SharePoint Get items operation paginates beyond the site's expected item count.

## Record results

Record deployment versions, tenant/site, test accounts, date, cases exercised, observed revocation delay, and any unresolved failure. Do not treat this checklist as passed until the checks have actually run.
