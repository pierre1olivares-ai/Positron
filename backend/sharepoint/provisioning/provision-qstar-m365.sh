#!/usr/bin/env bash
# =====================================================================
#  Q-Star Issue Manager — production provisioning (CLI for Microsoft 365)
# =====================================================================
#  Same schema as provision-qstar.ps1, built with the cross-platform
#  CLI for Microsoft 365 (no PowerShell required). Field XML is used so
#  the exact internal names are honoured (AddFieldInternalNameHint).
#
#  PREREQUISITES
#    npm i -g @pnp/cli-microsoft365
#    m365 login
#
#  RUN
#    SITE="https://contoso.sharepoint.com/sites/Quality" ./provision-qstar-m365.sh
#    # set PERSON_AS_TEXT=1 to store owners as text + email columns.
#    # For a beta without groups, run provision-qstar-beta-m365.sh instead.
# =====================================================================
set -euo pipefail

SITE="${SITE:?Set SITE to your site URL, e.g. https://contoso.sharepoint.com/sites/Quality}"
ISSUES="${ISSUES:-Q-Star Issues}"
PROGRESS="${PROGRESS:-Q-Star Progress Log}"
CONFIG="${CONFIG:-Q-Star Config}"
PERSON_AS_TEXT="${PERSON_AS_TEXT:-0}"
CREATE_ROLE_GROUPS="${CREATE_ROLE_GROUPS:-1}"

ADMIN_GROUP="${ADMIN_GROUP:-Q-Star Admins}"
QM_GROUP="${QM_GROUP:-Q-Star Quality Managers}"
OWNER_GROUP="${OWNER_GROUP:-Q-Star Task Owners}"
READER_GROUP="${READER_GROUP:-Q-Star Readers}"

# Built-in SharePoint permission-level IDs.
ROLE_FULL_CONTROL=1073741829
ROLE_EDIT=1073741830
ROLE_READ=1073741826

xml_escape() { printf '%s' "$1" | sed -e 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

# Build <CHOICES>..</CHOICES> from args (bash 3.2-safe; preserves spaces).
mk_choices() {
  local out="<CHOICES>" c
  for c in "$@"; do out="${out}<CHOICE>$(xml_escape "$c")</CHOICE>"; done
  printf '%s</CHOICES>' "$out"
}

# ---------- choice sets (keep in sync with the app) ----------
SEVERITY_XML="$(mk_choices "Critical" "High" "Medium" "Low")"
STATUS_XML="$(mk_choices "Created" "In Progress" "Under Testing/Revision" "On Hold" "Closed" "Rejected")"
TRANSFORM_XML="$(mk_choices "OFI" "NC Minor" "NC Major" "Only sent to Dept/BU for Action")"   # add "REC" when adopted
DEVIATION_XML="$(mk_choices "Communication" "Compliance" "Documentation" "Equipment" "Process" "Quality" "Safety" "System")"
ORIGIN_XML="$(mk_choices "Customer Complaints or Claims" "Internal Finding")"
REGION_XML="$(mk_choices "Germany" "Americas" "Asia Pacific" "China" "Eastern Europe" "Head Office" "Western Europe")"
YESNO_XML="$(mk_choices "Yes" "No")"
BU_XML="$(mk_choices "BU Aftermarket" "BU Airlines" "BU Automotive" "BU Diplo & High Security" \
  "BU High Tech & SemiCon" "BU Life Science" "Central Europe & Commercial Services" \
  "Claims & Complaints" "Customer Solution & Business Development" \
  "Digital Transformation & Data Management" "Finance & Controlling" "Human Resources" \
  "IT" "Legal & Data Protection" "Marketing" "Network & Products" "Quality" \
  "Risk Management" "Strategy & Transformation" "tmCT FRA" "tmCT MUC" "tmCT MEX/NLU" "tmCT PVG")"

ensure_list() {
  local title="$1"
  if m365 spo list get --webUrl "$SITE" --title "$title" >/dev/null 2>&1; then
    echo "= list '$title' exists"
  else
    m365 spo list add --webUrl "$SITE" --title "$title" --baseTemplate GenericList >/dev/null
    echo "+ list '$title'"
  fi
}

ensure_group() {
  local name="$1" description="$2" role_id="$3"
  if m365 spo group get --webUrl "$SITE" --name "$name" >/dev/null 2>&1; then
    echo "= group '$name' exists"
  else
    m365 spo group add --webUrl "$SITE" --name "$name" --description "$description" >/dev/null
    echo "+ group '$name'"
  fi
  if m365 spo web roleassignment add --webUrl "$SITE" --groupName "$name" --roleDefinitionId "$role_id" >/dev/null 2>&1; then
    echo "  + site permission assigned"
  else
    echo "  = site permission exists/skip"
  fi
}

# add_field <list> <field-xml> [view]   ; view=1 also adds to default view
add_field() {
  local list="$1" xml="$2" view="${3:-0}" opts="AddFieldInternalNameHint"
  [ "$view" = "1" ] && opts="${opts},AddToDefaultView"
  local name; name="$(printf '%s' "$xml" | grep -o "Name='[^']*'" | head -1)"
  if m365 spo field add --webUrl "$SITE" --listTitle "$list" --xml "$xml" --options "$opts" >/dev/null 2>&1; then
    echo "  + $name"
  else
    echo "  = (exists/skip) $name"
  fi
}

# typed field builders: <list> <internal> <display> [view] [required]
f_number()  { add_field "$1" "<Field Type='Number' Name='$2' DisplayName='$3' Required='${5:-FALSE}' Indexed='${6:-FALSE}'/>" "${4:-0}"; }
f_text()    { add_field "$1" "<Field Type='Text' Name='$2' DisplayName='$3' Required='${5:-FALSE}'/>" "${4:-0}"; }
f_note()    { add_field "$1" "<Field Type='Note' Name='$2' DisplayName='$3' NumLines='6' RichText='FALSE' Required='${5:-FALSE}'/>" "${4:-0}"; }
f_dateonly(){ add_field "$1" "<Field Type='DateTime' Name='$2' DisplayName='$3' Format='DateOnly' Required='${5:-FALSE}' Indexed='${6:-FALSE}'/>" "${4:-0}"; }
f_datetime(){ add_field "$1" "<Field Type='DateTime' Name='$2' DisplayName='$3'/>" "${4:-0}"; }
f_user()    { add_field "$1" "<Field Type='User' Name='$2' DisplayName='$3'/>" "${4:-0}"; }

# f_choice <list> <internal> <display> <choices-xml> [view] [default] [required] [indexed]
f_choice() {
  local def=""
  [ -n "${6:-}" ] && def="<Default>$(xml_escape "$6")</Default>"
  add_field "$1" "<Field Type='Choice' Name='$2' DisplayName='$3' Required='${7:-FALSE}' Indexed='${8:-FALSE}'>${4}${def}</Field>" "${5:-0}"
}

# f_person <list> <internal> <display> [view] : proper User column, or Text + Email companion
f_person() {
  local list="$1" name="$2" disp="$3" view="${4:-0}"
  if [ "$PERSON_AS_TEXT" = "1" ]; then
    f_text "$list" "$name" "$disp" "$view"
    f_text "$list" "${name}Email" "$disp Email" 0
  else
    f_user "$list" "$name" "$disp" "$view"
  fi
}

echo "Provisioning on $SITE"

if [ "$CREATE_ROLE_GROUPS" = "1" ]; then
  echo "--- Q-Star role groups ---"
  ensure_group "$ADMIN_GROUP"  "Q-Star application administrators" "$ROLE_FULL_CONTROL"
  ensure_group "$QM_GROUP"     "Q-Star Quality Managers"           "$ROLE_EDIT"
  ensure_group "$OWNER_GROUP"  "Q-Star task owners"                "$ROLE_READ"
  ensure_group "$READER_GROUP" "Q-Star read-only users"            "$ROLE_READ"
else
  echo "--- Q-Star role groups skipped by beta entry point ---"
fi

# =====================================================================
#  Q-Star Issues
# =====================================================================
echo "--- $ISSUES ---"
ensure_list "$ISSUES"

# Intake
f_number   "$ISSUES" "QsNumber"           "Qs Number"               1 TRUE
f_text     "$ISSUES" "ShortSummary"       "Short Summary"           1 TRUE
f_note     "$ISSUES" "Description"         "Description"             0 TRUE
f_note     "$ISSUES" "ImmediateAction"    "Immediate Action taken"  0
f_choice   "$ISSUES" "Severity"           "Severity"                "$SEVERITY_XML"  1 "" TRUE
f_person   "$ISSUES" "ReportedBy"         "Reported By"             0
f_dateonly "$ISSUES" "ReportDate"         "Report date"             1
f_choice   "$ISSUES" "DepartmentBU"       "Department/Business Unit" "$BU_XML"       1 "" TRUE
f_choice   "$ISSUES" "Region"             "Region"                  "$REGION_XML"   0 "" TRUE
f_choice   "$ISSUES" "AlreadyInContact"   "Already in Contact"      "$YESNO_XML"
f_choice   "$ISSUES" "DeviationType"      "Deviation Type"          "$DEVIATION_XML"
f_choice   "$ISSUES" "Origin"             "Origin"                  "$ORIGIN_XML"
f_note     "$ISSUES" "AdditionalComments" "Additional Comments"     0

# QM assessment
f_note     "$ISSUES" "FollowUp"           "Follow up"               0
f_choice   "$ISSUES" "Status"             "Status"                  "$STATUS_XML"    1 "" FALSE TRUE
f_choice   "$ISSUES" "TransformedInto"    "Transformed into"        "$TRANSFORM_XML"
f_choice   "$ISSUES" "TaskCreated"        "Task Created"            "$YESNO_XML"     0 "No"

# New fields (owner assignment, escalation, §10.2, NC effectiveness test)
f_choice   "$ISSUES" "Triaged"            "Triaged"                 "$YESNO_XML"     1 "No" FALSE TRUE
f_person   "$ISSUES" "TaskOwner"          "Task Owner"              1
f_text     "$ISSUES" "PermissionedOwnerEmail" "Permissioned Owner Email" 0
f_choice   "$ISSUES" "EscalationBU"       "Escalation BU"           "$BU_XML"
f_dateonly "$ISSUES" "DueDate"            "Due Date"                1 FALSE TRUE
f_note     "$ISSUES" "RootCause"          "Root Cause"              0
f_note     "$ISSUES" "CorrectiveAction"   "Corrective Action"       0
f_dateonly "$ISSUES" "ImplementationDate" "Implementation Date"     0
f_note     "$ISSUES" "EffectivenessCheck" "Effectiveness Check"     0
f_person   "$ISSUES" "VerifiedBy"         "Verified By"             0
f_dateonly "$ISSUES" "VerifiedDate"       "Verified Date"           0
f_dateonly "$ISSUES" "ClosedDate"         "Closed Date"             0
f_datetime "$ISSUES" "ClosedAt"          "Closed At"               0
f_note     "$ISSUES" "HoldReason"        "Hold Reason"             0
f_dateonly "$ISSUES" "HoldUntil"         "Hold Until"              0
f_choice   "$ISSUES" "OwnerUpdate"       "Owner Update"            "$YESNO_XML" 0 "No"
f_datetime "$ISSUES" "OwnerUpdateAt"     "Owner Update At"         0
f_note     "$ISSUES" "OwnerUpdateText"  "Owner Update Text"       0

echo "  Indexes requested for Status, Triaged and DueDate. Verify them in List settings after provisioning."

# =====================================================================
#  Q-Star Progress Log (append-only child list)
# =====================================================================
echo "--- $PROGRESS ---"
ensure_list "$PROGRESS"
f_number   "$PROGRESS" "ParentItemId" "Parent Item Id" 1 TRUE TRUE
f_person   "$PROGRESS" "Author"       "Author"         1
f_datetime "$PROGRESS" "EntryDate"    "Entry Date"     1
f_note     "$PROGRESS" "EntryText"    "Text"           0 TRUE
echo "  (append-only: only ever create items in this list — never edit)"

# =====================================================================
#  Q-Star Config (single-item settings store for the IT-settings tab)
# =====================================================================
echo "--- $CONFIG ---"
ensure_list "$CONFIG"
f_note     "$CONFIG" "SettingsJson" "Settings JSON" 0

echo "Done."
if [ "$CREATE_ROLE_GROUPS" = "1" ]; then
  echo "Production profile complete: four Q-Star role groups and site permissions were provisioned."
  echo "Next: add users or Entra security-group principals to the Q-Star groups."
  echo "Configure the assignment flow to grant/revoke item Edit permission for the assigned Task Owner."
else
  echo "Beta profile complete: no Q-Star groups or role assignments were created."
  echo "Next: enable Beta access mode on the web part and use the site's existing Owners, Members, and Visitors."
fi
