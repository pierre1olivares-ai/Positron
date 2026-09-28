<#
=====================================================================
 Q-Star Issue Manager — production provisioning (PnP PowerShell)
=====================================================================
 Creates the Q-Star lists with every column, internal name, choice value
 and index used by the app, plus the four production role groups and their
 site permissions. Use provision-qstar-beta.ps1 for a beta without groups.

 PREREQUISITES
   1. Install the module:      Install-Module PnP.PowerShell -Scope CurrentUser
   2. Register an Entra app for PnP (one-off, tenant admin):
          Register-PnPEntraIDApp -ApplicationName "PnP-QStar" -Tenant contoso.onmwo... -Interactive
      then pass its client id as -ClientId below.

 RUN
   ./provision-qstar.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Quality" -ClientId "<app-guid>"
   # add -PersonAsText to store owners as text + email columns instead of
   # proper Person columns (matches the test build's string model exactly).
=====================================================================
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$SiteUrl,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [string]$IssuesList   = "Q-Star Issues",
  [string]$ProgressList = "Q-Star Progress Log",
  [string]$ConfigList   = "Q-Star Config",
  [switch]$PersonAsText,
  [switch]$SkipRoleGroups
)

$ErrorActionPreference = "Stop"

# ---------- Choice value sets (must match the app) ----------
$SEVERITY   = "Critical","High","Medium","Low"
$STATUS     = "Created","In Progress","Under Testing/Revision","On Hold","Closed","Rejected"
$TRANSFORM  = "OFI","NC Minor","NC Major","Only sent to Dept/BU for Action"   # add "REC" when adopted
$DEVIATION  = "Communication","Compliance","Documentation","Equipment","Process","Quality","Safety","System"
$ORIGIN     = "Customer Complaints or Claims","Internal Finding"
$REGION     = "Americas (Miami)","Asia Pacific","China (Shanghai)","Eastern Europe (Vienna)","Head Office (Neu-Isenburg)","Western Europe (Amsterdam)"
$YESNO      = "Yes","No"
$BU = @(
  "BU Aftermarket","BU Airlines","BU Automotive","BU Diplo & High Security",
  "BU High Tech & SemiCon","BU Life Science","Central Europe & Commercial Services",
  "Claims & Complaints","Customer Solution & Business Development",
  "Digital Transformation & Data Management","Finance & Controlling","Human Resources",
  "IT","Legal & Data Protection","Marketing","Network & Products","Quality",
  "Risk Management","Strategy & Transformation","tmCT FRA","tmCT MUC","tmCT MEX/NLU","tmCT PVG"
)  # extend to the full set used on your live form

# ---------- Connect ----------
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ---------- Role groups ----------
function Ensure-RoleGroup {
  param([string]$Name,[string]$Description,[string]$Role)
  $group = Get-PnPGroup -Identity $Name -ErrorAction SilentlyContinue
  if (-not $group) {
    $group = New-PnPGroup -Title $Name -Description $Description
    Write-Host "+ group '$Name'" -ForegroundColor Green
  } else {
    Write-Host "= group '$Name' exists" -ForegroundColor DarkGray
  }
  try { Set-PnPGroupPermissions -Identity $Name -AddRole $Role | Out-Null }
  catch { Write-Host "  = permission '$Role' exists/skip" -ForegroundColor DarkGray }
}

if (-not $SkipRoleGroups) {
  Write-Host "`n--- Q-Star role groups ---" -ForegroundColor Cyan
  Ensure-RoleGroup -Name "Q-Star Admins" -Description "Q-Star application administrators" -Role "Full Control"
  Ensure-RoleGroup -Name "Q-Star Quality Managers" -Description "Q-Star Quality Managers" -Role "Edit"
  Ensure-RoleGroup -Name "Q-Star Task Owners" -Description "Q-Star task owners" -Role "Read"
  Ensure-RoleGroup -Name "Q-Star Readers" -Description "Q-Star read-only users" -Role "Read"
} else {
  Write-Host "`n--- Q-Star role groups skipped by beta entry point ---" -ForegroundColor Yellow
}

# ---------- Helpers ----------
function Ensure-List {
  param([string]$Title)
  if (Get-PnPList -Identity $Title -ErrorAction SilentlyContinue) {
    Write-Host "= list '$Title' exists" -ForegroundColor DarkGray
  } else {
    New-PnPList -Title $Title -Template GenericList -OnQuickLaunch | Out-Null
    Write-Host "+ list '$Title'" -ForegroundColor Green
  }
  # Title column is unused as a headline here — make it optional.
  try { Set-PnPField -List $Title -Identity "Title" -Values @{ Required = $false } -ErrorAction SilentlyContinue | Out-Null } catch {}
}

function Ensure-Field {
  param(
    [string]$List,[string]$Display,[string]$Internal,[string]$Type,
    [string[]]$Choices,[switch]$DateOnly,[switch]$Required,[switch]$AddToView,
    [string]$Default,[switch]$Index
  )

  # Person handling: proper User column, or text + companion email column.
  if ($Type -eq "Person") {
    if ($PersonAsText) {
      Ensure-Field -List $List -Display $Display -Internal $Internal -Type "Text" -AddToView:$AddToView
      Ensure-Field -List $List -Display "$Display Email" -Internal "${Internal}Email" -Type "Text"
      return
    } else { $Type = "User" }
  }

  if (Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue) {
    Write-Host "  = $Internal" -ForegroundColor DarkGray
  } else {
    $p = @{ List = $List; DisplayName = $Display; InternalName = $Internal; Type = $Type }
    if ($Choices)   { $p.Choices = $Choices }
    if ($Required)  { $p.Required = $true }
    if ($AddToView) { $p.AddToDefaultView = $true }
    Add-PnPField @p | Out-Null
    Write-Host "  + $Internal ($Type)" -ForegroundColor Green
  }

  # Post-create tweaks
  $vals = @{}
  if ($Type -eq "Note")             { $vals.RichText = $false }      # plain multiline
  if ($DateOnly)                    { $vals.DisplayFormat = 0 }      # 0 = DateOnly
  if ($PSBoundParameters.ContainsKey("Default")) { $vals.DefaultValue = $Default }
  if ($Index)                       { $vals.Indexed = $true }
  if ($vals.Count -gt 0) { try { Set-PnPField -List $List -Identity $Internal -Values $vals | Out-Null } catch { Write-Warning "  ! could not set props on $Internal: $($_.Exception.Message)" } }
}

# =====================================================================
#  Q-Star Issues
# =====================================================================
Write-Host "`n--- $IssuesList ---" -ForegroundColor Cyan
Ensure-List -Title $IssuesList

# Intake fields
Ensure-Field -List $IssuesList -Display "Qs Number"               -Internal "QsNumber"          -Type Number   -Required -AddToView
Ensure-Field -List $IssuesList -Display "Short Summary"           -Internal "ShortSummary"      -Type Text     -Required -AddToView
Ensure-Field -List $IssuesList -Display "Description"             -Internal "Description"       -Type Note     -Required
Ensure-Field -List $IssuesList -Display "Immediate Action taken"  -Internal "ImmediateAction"   -Type Note
Ensure-Field -List $IssuesList -Display "Severity"                -Internal "Severity"          -Type Choice -Choices $SEVERITY  -Required -AddToView
Ensure-Field -List $IssuesList -Display "Reported By"             -Internal "ReportedBy"        -Type Person
Ensure-Field -List $IssuesList -Display "Report date"            -Internal "ReportDate"        -Type DateTime -DateOnly -AddToView
Ensure-Field -List $IssuesList -Display "Department/Business Unit" -Internal "DepartmentBU"     -Type Choice -Choices $BU        -Required -AddToView
Ensure-Field -List $IssuesList -Display "Region"                  -Internal "Region"            -Type Choice -Choices $REGION    -Required
Ensure-Field -List $IssuesList -Display "Already in Contact"      -Internal "AlreadyInContact"  -Type Choice -Choices $YESNO
Ensure-Field -List $IssuesList -Display "Deviation Type"          -Internal "DeviationType"     -Type Choice -Choices $DEVIATION
Ensure-Field -List $IssuesList -Display "Origin"                  -Internal "Origin"            -Type Choice -Choices $ORIGIN
Ensure-Field -List $IssuesList -Display "Additional Comments"     -Internal "AdditionalComments" -Type Note

# QM assessment fields
Ensure-Field -List $IssuesList -Display "Follow up"               -Internal "FollowUp"          -Type Note
Ensure-Field -List $IssuesList -Display "Status"                  -Internal "Status"            -Type Choice -Choices $STATUS -AddToView -Index
Ensure-Field -List $IssuesList -Display "Transformed into"        -Internal "TransformedInto"   -Type Choice -Choices $TRANSFORM
Ensure-Field -List $IssuesList -Display "Task Created"            -Internal "TaskCreated"       -Type Choice -Choices $YESNO -Default "No"

# New fields (owner assignment, escalation, §10.2, NC effectiveness test)
Ensure-Field -List $IssuesList -Display "Triaged"                -Internal "Triaged"           -Type Choice -Choices $YESNO -Default "No" -AddToView -Index
Ensure-Field -List $IssuesList -Display "Task Owner"             -Internal "TaskOwner"         -Type Person -AddToView
Ensure-Field -List $IssuesList -Display "Permissioned Owner Email" -Internal "PermissionedOwnerEmail" -Type Text
Ensure-Field -List $IssuesList -Display "Escalation BU"          -Internal "EscalationBU"      -Type Choice -Choices $BU
Ensure-Field -List $IssuesList -Display "Due Date"               -Internal "DueDate"           -Type DateTime -DateOnly -AddToView -Index
Ensure-Field -List $IssuesList -Display "Root Cause"             -Internal "RootCause"         -Type Note
Ensure-Field -List $IssuesList -Display "Corrective Action"      -Internal "CorrectiveAction"  -Type Note
Ensure-Field -List $IssuesList -Display "Implementation Date"    -Internal "ImplementationDate" -Type DateTime -DateOnly
Ensure-Field -List $IssuesList -Display "Effectiveness Check"    -Internal "EffectivenessCheck" -Type Note
Ensure-Field -List $IssuesList -Display "Verified By"            -Internal "VerifiedBy"        -Type Person
Ensure-Field -List $IssuesList -Display "Verified Date"          -Internal "VerifiedDate"      -Type DateTime -DateOnly
Ensure-Field -List $IssuesList -Display "Closed Date"           -Internal "ClosedDate"        -Type DateTime -DateOnly
Ensure-Field -List $IssuesList -Display "Closed At"             -Internal "ClosedAt"          -Type DateTime
Ensure-Field -List $IssuesList -Display "Hold Reason"           -Internal "HoldReason"        -Type Note
Ensure-Field -List $IssuesList -Display "Hold Until"            -Internal "HoldUntil"         -Type DateTime -DateOnly
Ensure-Field -List $IssuesList -Display "Owner Update"          -Internal "OwnerUpdate"       -Type Choice -Choices $YESNO -Default "No"
Ensure-Field -List $IssuesList -Display "Owner Update At"       -Internal "OwnerUpdateAt"     -Type DateTime
Ensure-Field -List $IssuesList -Display "Owner Update Text"     -Internal "OwnerUpdateText"   -Type Note

# Make sure Status has the new "Under Testing/Revision" value even if the list pre-existed
try { Set-PnPField -List $IssuesList -Identity "Status" -Values @{ Choices = [string[]]$STATUS } | Out-Null }
catch { Write-Warning "Could not update Status choices automatically — add 'Under Testing/Revision' manually if missing." }

# =====================================================================
#  Q-Star Progress Log (append-only child list)
# =====================================================================
Write-Host "`n--- $ProgressList ---" -ForegroundColor Cyan
Ensure-List -Title $ProgressList
Ensure-Field -List $ProgressList -Display "Parent Item Id" -Internal "ParentItemId" -Type Number   -Required -AddToView -Index
Ensure-Field -List $ProgressList -Display "Author"         -Internal "Author"       -Type Person   -AddToView
Ensure-Field -List $ProgressList -Display "Entry Date"     -Internal "EntryDate"    -Type DateTime -AddToView
Ensure-Field -List $ProgressList -Display "Text"           -Internal "EntryText"    -Type Note     -Required
Write-Host "  (enforce append-only by only ever creating items in this list — never edit them)" -ForegroundColor DarkGray

# =====================================================================
#  Q-Star Config (single-item settings store for the IT-settings tab)
# =====================================================================
Write-Host "`n--- $ConfigList ---" -ForegroundColor Cyan
Ensure-List -Title $ConfigList
Ensure-Field -List $ConfigList -Display "Settings JSON" -Internal "SettingsJson" -Type Note

Write-Host "`nDone. Lists provisioned on $SiteUrl." -ForegroundColor Green
if ($SkipRoleGroups) {
  Write-Host "Beta profile complete: no Q-Star groups or role assignments were created." -ForegroundColor Yellow
  Write-Host "Next: enable Beta access mode on the web part and use the site's existing Owners, Members, and Visitors."
} else {
  Write-Host "Production profile complete: four Q-Star role groups and site permissions were provisioned." -ForegroundColor Green
  Write-Host "Next: add users or Entra groups to the Q-Star groups and configure the assignment-permission, intake, and reminder flows."
}
