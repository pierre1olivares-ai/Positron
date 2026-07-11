# Connection test plan — SharePoint & Power Automate

This is a manual verification checklist for confirming the Q-Star system works end-to-end in your company's Microsoft 365 tenant. It exists because none of this can be exercised from outside your tenant: it needs a real signed-in user, a real SharePoint site, and Power Automate running under your tenant's connections. Run through it once after provisioning, and again after any schema or flow change.

## 0. Prerequisites

- [ ] `backend/sharepoint/provisioning/provision-qstar.ps1` (or `provision-qstar-m365.sh`) has been run against your Quality site, with `-PersonAsText` (or `PERSON_AS_TEXT=1`).
- [ ] The web part builds locally (`cd frontend && npm install && npx gulp bundle`) with no errors.
- [ ] You have edit access to the Quality site to add the web part to a test page.

## 1. SharePoint connection (automated self-test, built into the app)

1. Add the Q-Star Issue Manager web part to a test page on the Quality site (via `gulp serve` pointed at your tenant's Workbench, or deploy the `.sppkg` to the app catalog first).
2. Open the Admin / IT-settings tab and click **Run Connection Test**.
3. Confirm every check passes:
   - Site access
   - Signed-in user resolved
   - `Q-Star Issues` list found, all expected columns present
   - `Q-Star Progress Log` list found, all expected columns present
   - `Under Testing/Revision` present in the Status choice list
   - `Status` / `Triaged` / `DueDate` indexed (warn is OK, fail is not)
   - Round-trip write test (create → update → delete) passes

If anything fails, the message names the missing list/column — re-run the provisioning script (it's idempotent) and retry.

## 2. Manual SharePoint sanity check (if you don't have the web part running yet)

1. Open `Q-Star Issues` in SharePoint (list view). Confirm you can add an item and edit `Status` to `Under Testing/Revision` without SharePoint rejecting the choice value.
2. Open `Q-Star Progress Log`. Confirm you can add an item with a `Parent Item Id` pointing at a real issue's ID.
3. Confirm your account, and one account per role (Admin / Quality Manager / Task Owner / Reader), can read the list — this previews permission issues before the app's role gating goes live.

## 3. Intake flow (Microsoft Form → list item)

1. Submit the Q-Star Microsoft Form as a test user.
2. In Power Automate, open the intake flow's **run history** and confirm the run succeeded.
3. In `Q-Star Issues`, confirm a new item appeared with `Triaged = No`, `Task Created = No`, `Status` empty, and a `QsNumber` one higher than the previous max.
4. Delete the test item afterward so it doesn't pollute the triage queue.

## 4. Daily reminder flow — seeded scenarios

Create three throwaway items (delete them after) to exercise each branch in `qstar-sharepoint-graph-integration.md` §5:

| Scenario | Setup | Expected on next flow run |
|---|---|---|
| **Overdue task** (rules 3–5) | `Status = In Progress`, `TaskCreated = Yes`, `DueDate` = 10 days ago, `TaskOwner`/`EscalationBU` set | Task Owner gets an overdue nudge; Quality Team gets a QM alert; BU lead gets an escalation (since overdue ≥ 7 days) |
| **NC mid-test** (rule 6) | `TransformedInto = NC Major`, `Status = Under Testing/Revision`, `ImplementationDate` = 55 days ago (test ends in ~5 days) | Task Owner gets a "test ending soon" notice, **not** an overdue nudge |
| **NC test elapsed** (rule 7) | Same as above but `ImplementationDate` = 61+ days ago | Task Owner **and** Quality Team get a "verify & close" notice |

For each: trigger the flow manually (**Run flow** > **Run now** in the Power Automate designer, or wait for the daily schedule), then check the run history for the expected branch and recipient, and confirm the actual email/Teams message arrived.

## 5. De-duplication

Run the reminder flow twice in the same day against the same overdue item. Confirm only one notification is sent (check whichever de-dupe mechanism was built — Reminder Log list or last-notified fields per §5.5) — the second run should skip it.

## 6. Sign-off

Once sections 1–5 pass with no unexpected failures, the connection between the web part, the SharePoint lists, and the Power Automate flows is confirmed working in your tenant. Record the date and who ran it here:

- Date:
- Run by:
- Notes:
