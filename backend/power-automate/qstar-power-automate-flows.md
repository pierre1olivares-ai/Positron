# Q-Star Issue Manager — Power Automate build guide

Three flows complete the production system:

- **Flow A — Intake:** copies each Q-Star Microsoft Form response into the `Q-Star Issues` list (Phase 1).
- **Flow B — Daily reminders:** the scheduled engine that reproduces the app's `buildReminders()` logic — owner nudges, QM alerts, BU escalation, and the NC effectiveness-test reminders.
- **Flow C — Assignment permissions:** reapplies item-level permissions whenever a Task Owner is assigned or changed.

Build them in the [Power Automate designer](https://make.powerautomate.com) on the same environment/site as the list. Field names below are the internal names created by the provisioning scripts.

> One-time prerequisite for Flow B's dedupe: create a small list **`Q-Star Reminder Log`** with columns `IssueId` (Number), `RuleKey` (Single line of text), `SentDate` (Single line of text, `yyyy-MM-dd`). And an optional **`BU Leads`** list with `Title` = BU name and `LeadEmail` (text) to resolve escalation recipients.

---

## Flow A — Intake (Form → SharePoint)

**Goal:** new form response → one list item with `Status` empty, `Triaged = No`, `TaskCreated = No`.

1. **Create → Automated cloud flow.** Name it `Q-Star — Intake`. Trigger: **Microsoft Forms → When a new response is submitted**. Select your Q-Star form.
2. Add **Microsoft Forms → Get response details.** Form Id = your form; Response Id = *Response Id* from the trigger. This exposes each answer as dynamic content.
3. Add **SharePoint → Get items** (used to compute the next QS number).
   - Site Address = your Quality site; List Name = `Q-Star Issues`.
   - **Order By** = `QsNumber desc`; **Top Count** = `1`.
4. Add **Compose** named `NextQsNumber` with this expression (falls back to 1041 on an empty list — match your live starting number):
   ```
   if(empty(outputs('Get_items')?['body/value']), 1041, add(first(outputs('Get_items')?['body/value'])?['QsNumber'], 1))
   ```
5. Add **SharePoint → Create item.** Site/List as above. Map fields:

   | List field | Value |
   |---|---|
   | `Title` | the form *Short summary* (or any non-empty value — Title is unused) |
   | `QsNumber` | `Outputs('NextQsNumber')` |
   | `ShortSummary` | form *Short summary* |
   | `Description` | form *Description* |
   | `ImmediateAction` | form *Immediate action taken* |
   | `Severity Value` | form *Severity* |
   | `ReportDate` | `utcNow()` |
   | `DepartmentBU Value` | form *Department / Business Unit* |
   | `Region Value` | form *Region* |
   | `AlreadyInContact Value` | form *Already in contact* (or `No`) |
   | `DeviationType Value` | form *Deviation type* |
   | `Origin Value` | form *Origin* |
   | `AdditionalComments` | form *Additional comments* |
   | `Status Value` | *(leave empty)* |
   | `Triaged Value` | `No` |
   | `TaskCreated Value` | `No` |

   *Choice columns appear in Create item as `<Field> Value` — that's where the string goes.*

6. **Reporter identity.** If `ReportedBy` is a **Person** column, set its **Claims** to the responder's email (add *Collect responders' email* on the form, then map *Responders' Email*). If you used the **text** model, set `ReportedBy` = responder name and `ReportedByEmail` = responder email instead.
7. **Save** and submit a test response to confirm an item lands in the triage queue (`Triaged = No`).

---

## Flow B — Daily reminders (the engine)

**Goal:** once a day, walk every active task and send the same reminders the app simulates. The rule set:

| Key | Applies to | Fires when | Recipient |
|---|---|---|---|
| `OWNER_HEADSUP` | open, not under test | 3 days before due | Task Owner |
| `OWNER_DUE` | open, not under test | on the due date | Task Owner |
| `OWNER_OVERDUE_{n}` | open, not under test | every 3 days overdue (n = days, ≤ 30) | Task Owner |
| `QM_ALERT` | overdue | 1 day after due | Quality Team |
| `BU_ESCALATION` | overdue ≥ 7 days | 7 days after due | Escalation BU lead |
| `NC_TEST_SOON` | NC under test | 7 days before test end | Task Owner |
| `NC_TEST_DONE` | NC under test | on/after test end | Task Owner **and** Quality Team |

The decisive split: an item with `Status = Under Testing/Revision` is open **but not overdue**, so it skips the first five rules and uses the last two (mirrors `isOverdue()` and `inNCTest()` in the app).

### B.1 Trigger and setup

1. **Create → Scheduled cloud flow.** Name `Q-Star — Daily reminders`. **Recurrence**: every `1` `Day`. Open the recurrence advanced options and set a sensible hour (e.g. `7`) and your time zone.
2. Add **Initialize variable** `varQualityTeam` (String) = your team address, e.g. `quality@time-matters.com`.
3. Add **Initialize variable** `varToday` (String) = expression `formatDateTime(utcNow(),'yyyy-MM-dd')`.

### B.2 Get the active tasks

4. Add **SharePoint → Get items** named `Get_active`.
   - Site/List = your Quality site / `Q-Star Issues`.
   - **Filter Query**:
     ```
     Status ne 'Closed' and Status ne 'Rejected' and TaskCreated eq 'Yes' and Triaged eq 'Yes'
     ```
   - **Top Count** = `5000`; in *Settings* on the action turn on **Pagination**. (`Status`, `Triaged`, `DueDate` were indexed by the provisioning script, so this filter is efficient.)

### B.3 Loop and classify

5. Add **Apply to each** over `Get_active → value`. Everything below sits inside this loop. (Where steps reference the current item, use the dynamic content from `Get_active`.)
6. Inside the loop add a **Compose `daysUntilDue`** (whole days; negative = overdue):
   ```
   div(sub(ticks(formatDateTime(items('Apply_to_each')?['DueDate'],'yyyy-MM-dd')), ticks(variables('varToday'))), 864000000000)
   ```
7. Add a **Compose `ownerEmail`**:
   - Person column: `items('Apply_to_each')?['TaskOwner']?['Email']`
   - Text model: `items('Apply_to_each')?['TaskOwnerEmail']`
8. Add a **Condition `isUnderTest`**: left = `items('Apply_to_each')?['Status']?['Value']` (or `…?['Status']` for text), operator **is equal to**, right = `Under Testing/Revision`.

### B.4 The reusable "dispatch" pattern

You'll repeat this small pattern for each rule, so build it once and copy it. Given a **RuleKey** and a **recipient/message**:

1. **Get items** on `Q-Star Reminder Log`, Filter Query:
   ```
   IssueId eq <current item ID> and RuleKey eq '<RuleKey>'
   ```
2. **Condition**: `length(body('Get_log')?['value'])` **is equal to** `0` (i.e. not sent before).
   - **If yes** →
     a. **Office 365 Outlook → Send an email (V2)** (and/or **Teams → Post message**) to the recipient.
     b. **SharePoint → Create item** in `Q-Star Reminder Log`: `IssueId` = current item ID, `RuleKey` = the key, `SentDate` = `variables('varToday')`.

This guarantees each milestone notifies exactly once even if a daily run is missed or repeated.

### B.5 NO branch (not under test) — rules 1–5

Inside the **"If no"** side of `isUnderTest`, first add a **Condition** that `DueDate` is not empty (`empty(items('Apply_to_each')?['DueDate'])` **is equal to** `false`), then add the five dispatch blocks. Wrap each in a Condition on `daysUntilDue` using `outputs('daysUntilDue')`:

| RuleKey | Gate condition (Condition action / expression) | Recipient | Suggested subject |
|---|---|---|---|
| `OWNER_HEADSUP` | `daysUntilDue` **=** `3` | `ownerEmail` | Due in 3 days: QS-{QsNumber} |
| `OWNER_DUE` | `daysUntilDue` **=** `0` | `ownerEmail` | Due today: QS-{QsNumber} |
| `OWNER_OVERDUE_{n}` | `and(less(outputs('daysUntilDue'),0), lessOrEquals(mul(outputs('daysUntilDue'),-1),30), equals(mod(mul(outputs('daysUntilDue'),-1),3),0))` | `ownerEmail` | Overdue {n} days: QS-{QsNumber} |
| `QM_ALERT` | `lessOrEquals(outputs('daysUntilDue'),-1)` | `varQualityTeam` | Overdue task — owner {TaskOwner} |
| `BU_ESCALATION` | `lessOrEquals(outputs('daysUntilDue'),-7)` | BU lead (see note) | Escalation: QS-{QsNumber} overdue 7+ days |

Notes:
- For `OWNER_OVERDUE_{n}`, build the RuleKey dynamically so each nudge is distinct: set it to `concat('OWNER_OVERDUE_', string(mul(outputs('daysUntilDue'),-1)))`.
- `QM_ALERT` and `BU_ESCALATION` use `lessOrEquals` rather than exact equality so a missed run still fires once (the dedupe log keeps it to one send).
- **BU lead lookup:** before the escalation dispatch, add **Get items** on `BU Leads`, Filter Query `Title eq '<EscalationBU value>'`, and use `first(body('Get_BU')?['value'])?['LeadEmail']` as the recipient. Fall back to `varQualityTeam` if the list returns nothing.

### B.6 YES branch (under test) — rules 6–7

Inside the **"If yes"** side of `isUnderTest`:

1. **Compose `testEnd`**:
   ```
   formatDateTime(addToTime(items('Apply_to_each')?['ImplementationDate'],2,'Month'),'yyyy-MM-dd')
   ```
2. **Compose `daysUntilTestEnd`**:
   ```
   div(sub(ticks(outputs('testEnd')), ticks(variables('varToday'))), 864000000000)
   ```
3. Two dispatch blocks:

| RuleKey | Gate | Recipient | Subject |
|---|---|---|---|
| `NC_TEST_SOON` | `daysUntilTestEnd` **=** `7` | `ownerEmail` | Effectiveness test ends in 7 days: QS-{QsNumber} |
| `NC_TEST_DONE` | `lessOrEquals(outputs('daysUntilTestEnd'),0)` | `ownerEmail` **and** `varQualityTeam` | Test complete — verify & close: QS-{QsNumber} |

`NC_TEST_DONE` for two recipients: either send one email with both addresses in **To**, or use two dispatch blocks with keys `NC_TEST_DONE_OWNER` and `NC_TEST_DONE_QM` so each is logged separately.

### B.7 Message body template

A simple, reusable body (Outlook V2 supports HTML):

```
QS-@{items('Apply_to_each')?['QsNumber']} — @{items('Apply_to_each')?['ShortSummary']}
Owner: @{items('Apply_to_each')?['TaskOwner']?['DisplayName']}
Severity: @{items('Apply_to_each')?['Severity']?['Value']}   Status: @{items('Apply_to_each')?['Status']?['Value']}
Due: @{formatDateTime(items('Apply_to_each')?['DueDate'],'dd MMM yyyy')}

Open in Q-Star: https://<your-app-url>/?issue=@{items('Apply_to_each')?['ID']}
```

---

## Flow C — Assignment permissions

**Goal:** SharePoint—not React—enforces that Admins/QMs can edit every issue, Readers and unassigned Task Owners can read, and only the currently assigned Task Owner can edit that item.

Use a dedicated flow connection owned by an approved Q-Star service account with Full Control on the site.

1. Create an **Automated cloud flow** named `Q-Star — Assignment permissions` with SharePoint trigger **When an item is created or modified** on `Q-Star Issues`.
2. Add a trigger condition so the flow runs when `TaskOwner` is populated. During testing, also record the last permissioned owner in a small text column such as `PermissionedOwnerEmail`; skip the flow when that value already equals the current Task Owner email. This prevents unnecessary permission churn.
3. Resolve these principal IDs once with **Send an HTTP request to SharePoint** (`GET`):
   - `/_api/web/sitegroups/getbyname('Q-Star Admins')?$select=Id`
   - `/_api/web/sitegroups/getbyname('Q-Star Quality Managers')?$select=Id`
   - `/_api/web/sitegroups/getbyname('Q-Star Task Owners')?$select=Id`
   - `/_api/web/sitegroups/getbyname('Q-Star Readers')?$select=Id`
4. Break inheritance on the current item and clear old assignments (`POST`):
   ```text
   /_api/web/lists/getbytitle('Q-Star Issues')/items(<ID>)/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)
   ```
5. Add role assignments with `POST` requests to:
   ```text
   /_api/web/lists/getbytitle('Q-Star Issues')/items(<ID>)/roleassignments/addroleassignment(principalid=<PRINCIPAL_ID>,roledefid=<ROLE_ID>)
   ```
   Apply this matrix:

   | Principal | Permission | Built-in role ID |
   |---|---|---:|
   | Q-Star Admins | Full Control | `1073741829` |
   | Q-Star Quality Managers | Edit | `1073741830` |
   | Q-Star Task Owners | Read | `1073741826` |
   | Q-Star Readers | Read | `1073741826` |
   | Current `TaskOwner/Id` | Edit | `1073741830` |

6. Update `PermissionedOwnerEmail` to the current Task Owner email. If owner assignment is cleared, reapply only the four group permissions and leave no individual Edit grant.
7. Configure failure alerts. A permissions-flow failure is a security-relevant operational event and must not be silently ignored.

Breaking inheritance again with `clearSubscopes=true` removes the previous owner's direct grant, so reassignment revokes old access automatically. Monitor the number of unique permission scopes as the register grows; if the List approaches SharePoint governance limits, review the item-security design with the SharePoint administrator.

---

## Testing checklist

Use the seeded scenarios as your test fixtures (or replicate them as real items):

1. **Overdue task** → on the right days you should see `OWNER_DUE`, then `OWNER_OVERDUE_3/6/…`, `QM_ALERT` (from due+1), and `BU_ESCALATION` (from due+7), each logged once.
2. **NC mid-test** (implemented ~5 weeks ago) → no overdue chasing; `NC_TEST_SOON` fires 7 days before the end date.
3. **NC past test end** → `NC_TEST_DONE` fires for owner + QM; the app then unlocks *Verify & close*.
4. Re-run the flow manually the same day → no duplicate emails (the Reminder Log blocks repeats).
5. Assign an issue to Owner A → Owner A can edit it while another Task Owner can only read it.
6. Reassign it to Owner B → Owner A immediately loses Edit and Owner B gains it; Admin/QM retain Edit throughout.

**Tips:** while testing, temporarily point recipients at your own mailbox and shorten the recurrence to hourly; switch back before go-live. To validate date logic without waiting, set a test item's `DueDate`/`ImplementationDate` so the target day equals today. Turn on **Pagination** on every `Get items` action that could exceed 100 rows.

---

*The rule keys, day offsets, the 2-month NC test window and the under-test branch are taken directly from the app's reminder engine, so the flow and the in-app "Reminders" panel stay in agreement.*
