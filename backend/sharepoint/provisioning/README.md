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

## Local checks

`node --test backend/sharepoint/provisioning/tests/provisioning.test.mjs` executes
the real Bash entry point against a stateful m365 test double; it never contacts a
tenant. It covers fresh provisioning, existing-list migration, a read-only
preview, offset stability, and failures. Live permission checks still require the
tenant test plan.
