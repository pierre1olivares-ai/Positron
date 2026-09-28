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
   # -RegionMigration Preview inspects legacy rows without writes.
   # -RegionMigration Apply migrates them during a maintenance window.
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
  [switch]$SkipRoleGroups,
  [ValidateSet("Preserve","Preview","Apply")][string]$RegionMigration = "Preserve",
  [ValidateSet("Preserve","Preview","Apply")][string]$ProgressMigration = "Preserve"
)

$ErrorActionPreference = "Stop"
if ($PersonAsText) { throw "PersonAsText is incompatible with the web part. Use native Person columns." }
$regionSchema = Get-Content "$PSScriptRoot/region-schema.json" -Raw | ConvertFrom-Json -AsHashtable

# ---------- Choice value sets (must match the app) ----------
$SEVERITY   = "Critical","High","Medium","Low"
$STATUS     = "Created","In Progress","Under Testing/Revision","On Hold","Closed","Rejected"
$TRANSFORM  = "OFI","NC Minor","NC Major","Only sent to Dept/BU for Action"   # add "REC" when adopted
$DEVIATION  = "Communication","Compliance","Documentation","Equipment","Process","Quality","Safety","System"
$ORIGIN     = "Customer Complaints or Claims","Internal Finding"
$REGION     = [string[]]$regionSchema.choices
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

function Sync-Regions {
  param([switch]$Preview)
  $rows = @(Get-PnPListItem -List $IssuesList -Fields "Region" -PageSize 1000)
  $changes = @($rows | Where-Object { $_["Region"] -and $regionSchema.aliases.ContainsKey([string]$_["Region"]) })
  foreach ($row in $changes) { Write-Host "Region item $($row.Id): $($row['Region']) -> $($regionSchema.aliases[[string]$row['Region']])" }
  if ($Preview) { Write-Host "Preview only: $($changes.Count) region changes; no writes performed."; return }
  $field = Get-PnPField -List $IssuesList -Identity "Region"
  $originalChoices = @(Get-PnPProperty -ClientObject $field -Property Choices)
  $choices = @($REGION) + $originalChoices + @($rows | ForEach-Object { $_["Region"] } | Where-Object { $_ })
  Set-PnPField -List $IssuesList -Identity "Region" -Values @{ Choices = [string[]]@($choices | Select-Object -Unique) } | Out-Null
  if ($RegionMigration -eq "Apply") {
    foreach ($row in $changes) {
      Set-PnPListItem -List $IssuesList -Identity $row.Id -Values @{ Region = $regionSchema.aliases[[string]$row["Region"]] } -UpdateType SystemUpdate | Out-Null
    }
    $customChoices = @($choices | Where-Object { -not $regionSchema.aliases.ContainsKey([string]$_) })
    Set-PnPField -List $IssuesList -Identity "Region" -Values @{ Choices = [string[]]@($customChoices | Select-Object -Unique) } | Out-Null
  } elseif ($changes.Count) {
    Write-Host "Legacy regions preserved. Review -RegionMigration Preview, then use Apply during maintenance."
  }
}
if ($RegionMigration -eq "Preview") { Sync-Regions -Preview; return }
$securityParameters = @{ SiteUrl = $SiteUrl; IssuesList = $IssuesList; ProgressList = $ProgressList; ConfigList = $ConfigList; Production = (-not $SkipRoleGroups); ProgressMigration = $ProgressMigration }
if ($ProgressMigration -eq "Preview") { & "$PSScriptRoot/secure-qstar.ps1" @securityParameters; return }
$progressPreflightDone = $false
if (Get-PnPList -Identity $ProgressList -ErrorAction SilentlyContinue) {
  & "$PSScriptRoot/secure-qstar.ps1" @securityParameters -Preflight
  $progressPreflightDone = $true
}

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
  Set-PnPGroupPermissions -Identity $Name -AddRole $Role | Out-Null
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
  Set-PnPField -List $Title -Identity "Title" -Values @{ Required = $false } | Out-Null
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

  $field = Get-PnPField -List $List -Identity $Internal -ErrorAction SilentlyContinue
  if ($field) {
    if ($field.TypeAsString -ne $Type) { throw "Field $List/$Internal is $($field.TypeAsString); expected $Type. Migrate it first; no data was converted." }
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
  if ($Choices) {
    $oldChoices = if ($field) { @(Get-PnPProperty -ClientObject $field -Property Choices) } else { @() }
    $vals.Choices = [string[]]@((@($Choices) + $oldChoices) | Select-Object -Unique)
  }
  if ($PSBoundParameters.ContainsKey("Required")) { $vals.Required = [bool]$Required }
  if ($Type -eq "Note")             { $vals.RichText = $false }      # plain multiline
  if ($DateOnly)                    { $vals.DisplayFormat = 0 }      # 0 = DateOnly
  if ($PSBoundParameters.ContainsKey("Default")) { $vals.DefaultValue = $Default }
  if ($Index)                       { $vals.Indexed = $true }
  if ($vals.Count -gt 0) { Set-PnPField -List $List -Identity $Internal -Values $vals | Out-Null }
}

# =====================================================================
#  Q-Star Issues
# =====================================================================
Write-Host "`n--- $IssuesList ---" -ForegroundColor Cyan
Ensure-List -Title $IssuesList

# Intake fields
Ensure-Field -List $IssuesList -Display "Qs Number"               -Internal "QsNumber"          -Type Number   -Required:$false -AddToView
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
Ensure-Field -List $IssuesList -Display "Reminder Cycle" -Internal "ReminderCycle" -Type Text
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

Sync-Regions

# =====================================================================
#  Q-Star Progress Log (append-only child list)
# =====================================================================
Write-Host "`n--- $ProgressList ---" -ForegroundColor Cyan
Ensure-List -Title $ProgressList
Ensure-Field -List $ProgressList -Display "Parent Item Id" -Internal "ParentItemId" -Type Number   -Required:$false -AddToView -Index
Ensure-Field -List $ProgressList -Display "Author"         -Internal "Author"       -Type Person   -AddToView
Ensure-Field -List $ProgressList -Display "Entry Date"     -Internal "EntryDate"    -Type DateTime -AddToView
Ensure-Field -List $ProgressList -Display "Text"           -Internal "EntryText"    -Type Note     -Required


# =====================================================================
#  Q-Star Config (single-item settings store for the IT-settings tab)
# =====================================================================
Write-Host "`n--- $ConfigList ---" -ForegroundColor Cyan
Ensure-List -Title $ConfigList
Ensure-Field -List $ConfigList -Display "Settings JSON" -Internal "SettingsJson" -Type Note
Ensure-Field -List $ConfigList -Display "Reference Offset" -Internal "ReferenceOffset" -Type Number
$settings = @(Get-PnPListItem -List $ConfigList -Fields "ReferenceOffset" -PageSize 1000)
if ($settings.Count -gt 1) { throw "$ConfigList must contain one settings item; reconcile duplicates before provisioning." }
if ($settings.Count -eq 1 -and $null -ne $settings[0]["ReferenceOffset"]) {
  $offset = [double]$settings[0]["ReferenceOffset"]
  if ($offset -lt 1000 -or $offset -ne [math]::Floor($offset)) { throw "Existing ReferenceOffset is invalid; it was not overwritten." }
  Write-Host "ReferenceOffset $offset preserved (immutable after first use)."
} else {
  $offset = 1000
  foreach ($row in @(Get-PnPListItem -List $IssuesList -Fields "QsNumber" -PageSize 1000)) {
    if ($null -eq $row["QsNumber"]) { continue }
    $number = [double]$row["QsNumber"]
    if ($number -lt 0 -or $number -ne [math]::Floor($number)) { throw "Invalid legacy QsNumber on item $($row.Id); resolve before allocating references." }
    $offset = [math]::Max($offset, $number)
  }
  if ($settings.Count) {
    Set-PnPListItem -List $ConfigList -Identity $settings[0].Id -Values @{ ReferenceOffset = $offset } | Out-Null
  } else {
    Add-PnPListItem -List $ConfigList -Values @{ Title = "Q-Star Settings"; ReferenceOffset = $offset } | Out-Null
  }
  Write-Host "ReferenceOffset initialized to $offset; existing QS references were not changed."
}

# Validate existing history before changing any permission assignments.
if (-not $progressPreflightDone) { & "$PSScriptRoot/secure-qstar.ps1" @securityParameters -Preflight }
if (-not $SkipRoleGroups) {
  Write-Host "`n--- Q-Star role groups ---" -ForegroundColor Cyan
  Ensure-RoleGroup -Name "Q-Star Admins" -Description "Q-Star application administrators" -Role "Full Control"
  Ensure-RoleGroup -Name "Q-Star Quality Managers" -Description "Q-Star Quality Managers" -Role "Edit"
  Ensure-RoleGroup -Name "Q-Star Task Owners" -Description "Q-Star task owners" -Role "Read"
  Ensure-RoleGroup -Name "Q-Star Readers" -Description "Q-Star read-only users" -Role "Read"
} else {
  Write-Host "`n--- Q-Star role groups skipped by beta entry point ---" -ForegroundColor Yellow
}

& "$PSScriptRoot/secure-qstar.ps1" @securityParameters

Write-Host "`nDone. Lists provisioned on $SiteUrl." -ForegroundColor Green
if ($SkipRoleGroups) {
  Write-Host "Beta profile complete: no Q-Star groups or role assignments were created." -ForegroundColor Yellow
  Write-Host "Next: enable Beta access mode on the web part and use the site's existing Owners, Members, and Visitors."
} else {
  Write-Host "Production profile complete: list, issue, and journal-folder permissions reconciled." -ForegroundColor Green
  Write-Host "Next: add users or Entra groups to the Q-Star groups and configure the assignment-permission, intake, and reminder flows."
}
