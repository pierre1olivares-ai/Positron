# Provisioning and upgrades

Use a maintenance window for an existing site: stop intake and reminder flows and
pause editing until the schema and data migration finish. Keep a list export and
version history. Scripts reconcile existing fields and indexes, stop on errors,
and never replace lists or overwrite existing QS references.

For beta use `provision-qstar-beta-m365.sh` or `provision-qstar-beta.ps1`. The Bash
entry point uses Node (already required by CLI for Microsoft 365). PowerShell uses
PnP.PowerShell. Both require native Person columns; the old PersonAsText mode is
rejected with an error instead of producing a schema the web part cannot read.
Existing text-person lists need a separately reviewed identity migration.

For production use `provision-qstar-m365.sh` or `provision-qstar.ps1`; these also
reconcile the permissions described below. Choose one toolchain for a run.
Run the migration examples from `backend/sharepoint/provisioning/`. Use the
corresponding beta entry point for a pilot: the migration options work in both
profiles. Beta preserves existing permissions; it does not undo production ACLs.

## Region migration

The mapping in `region-schema.json` follows the main-branch taxonomy: Germany is
part of Western Europe (Amsterdam); the other old region names map to their named
offices. Existing custom values are retained, never guessed or deleted.

Default provisioning (`Preserve`) adds the canonical choices while retaining
legacy choices and existing record values. Preview an existing site's row changes
without **any** writes:

```sh
SITE="https://contoso.sharepoint.com/sites/Quality" REGION_MIGRATION=preview bash provision-qstar-beta-m365.sh
```

```powershell
./provision-qstar-beta.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Quality" -ClientId "<app-guid>" -RegionMigration Preview
```

Use `REGION_MIGRATION=apply` / `-RegionMigration Apply` during maintenance to migrate
only the Region field of the displayed rows. SystemUpdate preserves Modified and
Editor. After successful migration the known legacy choices are retired; custom
choices remain. Repeating the migration is safe. A failed run can be resumed.

## Stable QS references

`QsNumber` is optional during creation. Provisioning creates a numeric
`ReferenceOffset` on the single `Q-Star Config` item, initialized once to at least
1000 and at least the maximum existing QsNumber. It preserves SettingsJson and all
existing nonempty QsNumber values. Multiple Config items or invalid numbers stop
the upgrade instead of choosing a value arbitrarily.

Every new producer must create the item first, then use
`ReferenceOffset + SharePoint ID`. Existing references remain unchanged. Keep the
offset immutable once intake resumes; raising it on each provisioning run would
change fallback references for items whose second write has not completed.
An already-populated offset is never recomputed. Deploy the updated web part and
intake flow together before restarting writes.

## Journal migration and production permissions

Progress entries now live in `issue-<SharePoint ID>` folders beneath the actual
Progress Log root URL. Folder location is the authoritative parent; writable
ParentItemId remains optional legacy metadata. Provisioning enables folders,
content types and versioning. Built-in Title is optional on all three lists.

For an existing register, inspect the journal before changing anything:

```sh
SITE="https://contoso.sharepoint.com/sites/Quality" PROGRESS_MIGRATION=preview bash provision-qstar-m365.sh
```

```powershell
./provision-qstar.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Quality" -ClientId "<app-guid>" -ProgressMigration Preview
```

Use `PROGRESS_MIGRATION=apply` / `-ProgressMigration Apply` during maintenance to
move valid loose rows in place. The script preserves item IDs and history and
verifies Author/Created after each move. Orphan parents and unrecognized folders
stop before writes; correct their mapping explicitly and repeat the preview.
Default Preserve refuses to upgrade loose rows without the reviewed Apply mode.
Region and progress previews are separate read-only runs. Both Apply modes may
be combined in the maintenance run. A partial migration can be resumed.

Production reconciles existing permissions, including removal of obsolete
owners and broad Members grants. It creates **Q-Star Append Progress** with Read
and Add Items, excluding Edit/Delete/Manage Permissions:

| Object | Admin | Quality Manager | Task Owners / Readers | Current individual owner |
|---|---|---|---|---|
| Issues root | Full | Edit | Read | — |
| Issue item | Full | Edit | Read | Edit |
| Progress root | Full | Append | Read | — |
| Issue progress folder | Full | Append | Read | Append |
| Config | Full | Read | Read | — |

The executing provisioning account retains Full to prevent lockout. Child
progress entries inherit folder permissions; production resets historical
entry-specific grants and any unique permissions on the singleton Config item. SharePoint-managed Limited Access on ancestors is retained.
Put the flow service identity in Q-Star Admins; site-level permissions alone do
not bypass these unique list ACLs. Admin maintenance access remains intentional.

Configure [Flow C](../../power-automate/qstar-power-automate-flows.md#c--reconcile-assignment-and-journal-permissions)
to create folders and reconcile both objects after each assignment, including an
empty owner. Permissions change asynchronously after the issue save. Beta uses
the same folder structure but retains site permissions, so it cannot demonstrate
production isolation or append-only enforcement. An owner without root Add
permission waits for Flow C to prepare their folder; QMs/Admins may lazily create
one only when their actual SharePoint permissions allow it.

Reminder event identity and the `ReminderCycle` fallback are defined in
[Flow B](../../power-automate/qstar-power-automate-flows.md#event-identity-and-delivery).

## Local checks

From the repository root, `node --test backend/sharepoint/provisioning/tests/*.test.mjs` executes
the real Bash entry point against a stateful m365 test double; it never contacts a
tenant. It covers fresh provisioning, existing-list migration, a read-only
preview, offset stability, and failures. Transport-level tests exercise ACL
reassignment, empty owners, role downgrades, paging, and in-place migration.
These are local contract tests, not tenant integration tests. PnP PowerShell and
the live permission checks still require the [tenant test plan](../connection-test-plan.md).
