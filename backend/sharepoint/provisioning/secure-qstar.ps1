# Reconcile existing and new lists. Requires the caller's active PnP connection.
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$SiteUrl,
  [string]$IssuesList = "Q-Star Issues",
  [string]$ProgressList = "Q-Star Progress Log",
  [string]$ConfigList = "Q-Star Config",
  [switch]$Production,
  [switch]$Preflight,
  [ValidateSet("Preserve","Preview","Apply")][string]$ProgressMigration = "Preserve"
)
$ErrorActionPreference = "Stop"
$FullRole = 1073741829; $EditRole = 1073741830; $ReadRole = 1073741826; $LimitedRole = 1073741825
$AppendName = "Q-Star Append Progress"
$AppendPermissions = @{ High = "48"; Low = "134418531" }
function Escape-QStarLiteral([string]$Value) { [uri]::EscapeDataString($Value.Replace("'", "''")).Replace("'", "%27") }
function Get-QStarListPath([string]$Title) { "web/lists/getbytitle('$(Escape-QStarLiteral $Title)')" }
function Invoke-QStarRest([string]$Method, [string]$Path, $Body) {
  $base = $SiteUrl.TrimEnd('/') + '/_api/'
  $url = if ($Path.StartsWith('https://')) { $Path } else { $base + $Path }
  if (-not $url.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { throw "Unexpected SharePoint pagination URL." }
  $args = @{ Method = $Method; Url = $url; ContentType = 'application/json;odata=nometadata' }
  if ($null -ne $Body) { $args.Content = $Body | ConvertTo-Json -Depth 20 -Compress }
  $response = Invoke-PnPSPRestMethod @args
  if ($response.d) { return $response.d }
  return $response
}
function Get-QStarAll([string]$Path) {
  do {
    $response = Invoke-QStarRest "Get" $Path
    if ($null -ne $response.value) { $response.value } elseif ($null -ne $response.results) { $response.results }
    $Path = $response.'@odata.nextLink'
    if (-not $Path) { $Path = $response.'odata.nextLink' }
    if (-not $Path) { $Path = $response.__next }
  } while ($Path)
}
function Sync-QStarAcl([string]$Path, [hashtable]$Desired) {
  $object = Invoke-QStarRest "Get" ($Path + '?$select=HasUniqueRoleAssignments')
  if (-not $object.HasUniqueRoleAssignments) { Invoke-QStarRest "Post" "$Path/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)" | Out-Null }
  $assignments = @(Get-QStarAll ($Path + '/roleassignments?$select=Member/Id,RoleDefinitionBindings/Id&$expand=Member,RoleDefinitionBindings'))
  $existing = @{}
  foreach ($assignment in $assignments) {
    $bindings = if ($null -ne $assignment.RoleDefinitionBindings.results) { @($assignment.RoleDefinitionBindings.results) } else { @($assignment.RoleDefinitionBindings) }
    $existing[[int]$assignment.Member.Id] = @($bindings | ForEach-Object { [int]$_.Id })
  }
  foreach ($principal in $Desired.Keys) {
    if ([int]$principal -le 0) { throw "Invalid permission principal ID." }
    foreach ($role in $Desired[$principal]) {
      if ($existing[$principal] -notcontains $role) { Invoke-QStarRest "Post" "$Path/roleassignments/addroleassignment(principalid=$principal,roledefid=$role)" | Out-Null }
    }
  }
  foreach ($principal in $existing.Keys) {
    foreach ($role in $existing[$principal]) {
      if ($role -ne $LimitedRole -and $Desired[$principal] -notcontains $role) { Invoke-QStarRest "Post" "$Path/roleassignments/removeroleassignment(principalid=$principal,roledefid=$role)" | Out-Null }
    }
  }
}
function New-QStarDesired([int]$QmRole, [int]$OwnerId = 0, [int]$OwnerRole = 0) {
  $desired = @{}
  $pairs = @(@($currentUser,$FullRole), @($groups[0],$FullRole), @($groups[1],$QmRole), @($groups[2],$ReadRole), @($groups[3],$ReadRole))
  if ($OwnerId -gt 0) { $pairs += ,@($OwnerId,$OwnerRole) }
  foreach ($pair in $pairs) { $principal = [int]$pair[0]; $desired[$principal] = @((@($desired[$principal]) + $pair[1]) | Where-Object { $null -ne $_ } | Select-Object -Unique) }
  return $desired
}
$issuesPath = Get-QStarListPath $IssuesList
$progressPath = Get-QStarListPath $ProgressList
$configPath = Get-QStarListPath $ConfigList
$progress = Invoke-QStarRest "Get" ($progressPath + '?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder')
$root = $progress.RootFolder.ServerRelativeUrl.TrimEnd('/')
$issues = @(Get-QStarAll ($issuesPath + '/items?$select=Id,TaskOwnerId&$top=2000'))
$issueIds = @{}; foreach ($issue in $issues) { $issueIds[[int]$issue.Id] = $true }
$entries = @(Get-QStarAll ($progressPath + '/items?$select=Id,FSObjType,ParentItemId,FileDirRef,FileRef,AuthorId,Created&$top=2000'))
$moves = @()
foreach ($entry in $entries | Where-Object { $_.FSObjType -ne 1 }) {
  if ($entry.FileDirRef -eq $root) {
    $parent = $entry.ParentItemId
    if ($null -eq $parent -or [double]$parent -ne [math]::Floor([double]$parent) -or -not $issueIds.ContainsKey([int]$parent)) { throw "Journal item $($entry.Id) has no valid existing parent; correct its mapping before migration. No history was discarded." }
    $moves += $entry
  } else {
    $prefix = $root + '/issue-'
    if (-not $entry.FileDirRef.StartsWith($prefix) -or $entry.FileDirRef.Substring($prefix.Length) -notmatch '^[1-9][0-9]*$' -or -not $issueIds.ContainsKey([int]$entry.FileDirRef.Substring($prefix.Length))) { throw "Journal item $($entry.Id) has an unrecognized folder. Review its parent mapping before migration." }
  }
}
foreach ($entry in $moves) { Write-Host "Journal item $($entry.Id): move to $root/issue-$($entry.ParentItemId)" }
if ($ProgressMigration -eq "Preview") { Write-Host "Preview only: no changes made."; return }
if ($moves.Count -and $ProgressMigration -ne "Apply") { throw "Legacy journal rows require a reviewed migration. Run -ProgressMigration Preview, then Apply during maintenance." }
if ($Preflight) { Write-Host "Journal migration preflight passed."; return }
$configItems = @()
if ($Production) {
  $configItems = @(Get-QStarAll ($configPath + '/items?$select=Id'))
  if ($configItems.Count -ne 1) { throw "Config must contain exactly one provisioned settings item before permissions are applied." }
}
$currentUser = [int](Invoke-QStarRest "Get" 'web/currentuser?$select=Id').Id
if ($Production) {
  $groups = @("Q-Star Admins", "Q-Star Quality Managers", "Q-Star Task Owners", "Q-Star Readers") | ForEach-Object { [int](Invoke-QStarRest "Get" "web/sitegroups/getbyname('$(Escape-QStarLiteral $_)')?`$select=Id").Id }
  $roles = @(Get-QStarAll 'web/roledefinitions?$select=Id,Name,BasePermissions')
  $append = $roles | Where-Object { $_.Name -eq $AppendName }
  if (-not $append) {
    Invoke-QStarRest "Post" 'web/roledefinitions' @{ Name = $AppendName; Description = 'Read and append progress; no edit, delete or permission management.'; BasePermissions = $AppendPermissions } | Out-Null
    $append = Invoke-QStarRest "Get" "web/roledefinitions/getbyname('$(Escape-QStarLiteral $AppendName)')"
  }
  if (-not $append.Id -or [string]$append.BasePermissions.High -ne $AppendPermissions.High -or [string]$append.BasePermissions.Low -ne $AppendPermissions.Low) { throw "Q-Star Append Progress has unexpected permissions. Review this permission level before proceeding." }
  $appendRole = [int]$append.Id
  Sync-QStarAcl $issuesPath (New-QStarDesired $EditRole)
  Sync-QStarAcl $progressPath (New-QStarDesired $appendRole)
  Sync-QStarAcl $configPath (New-QStarDesired $ReadRole)
  Invoke-QStarRest "Post" "$configPath/items($($configItems[0].Id))/resetroleinheritance()" | Out-Null
}
Set-PnPList -Identity $ProgressList -EnableContentTypes $true -EnableFolderCreation $true -EnableVersioning $true | Out-Null
Set-PnPList -Identity $IssuesList -EnableVersioning $true | Out-Null
$folderUrls = @{}; foreach ($folder in @(Get-QStarAll ($progressPath + '/RootFolder/Folders?$select=ServerRelativeUrl'))) { $folderUrls[$folder.ServerRelativeUrl] = $true }
foreach ($issue in $issues) {
  $path = "$root/issue-$($issue.Id)"
  if (-not $folderUrls.ContainsKey($path)) { Invoke-QStarRest "Post" "web/folders/addUsingPath(DecodedUrl='$(Escape-QStarLiteral $path)',overwrite=false)" | Out-Null }
  if ($Production) {
    Sync-QStarAcl "$issuesPath/items($($issue.Id))" (New-QStarDesired $EditRole ([int]$issue.TaskOwnerId) $EditRole)
    Sync-QStarAcl "web/GetFolderByServerRelativePath(decodedUrl='$(Escape-QStarLiteral $path)')/ListItemAllFields" (New-QStarDesired $appendRole ([int]$issue.TaskOwnerId) $appendRole)
  }
}
foreach ($entry in $moves) {
  $targetFolder = "$root/issue-$($entry.ParentItemId)"
  $target = $targetFolder + '/' + $entry.FileRef.Substring($entry.FileRef.LastIndexOf('/') + 1)
  $move = @{
    srcPath = @{ __metadata = @{ type = 'SP.ResourcePath' }; DecodedUrl = (([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority) + [string]$entry.FileRef) }
    destPath = @{ __metadata = @{ type = 'SP.ResourcePath' }; DecodedUrl = (([uri]$SiteUrl).GetLeftPart([System.UriPartial]::Authority) + $target) }
    options = @{ __metadata = @{ type = 'SP.MoveCopyOptions' }; KeepBoth = $false; ResetAuthorAndCreatedOnCopy = $false; RetainEditorAndModifiedOnMove = $true; ShouldBypassSharedLocks = $false }
  }
  Invoke-QStarRest "Post" 'SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=false' $move | Out-Null
  $moved = Invoke-QStarRest "Get" ($progressPath + "/items($($entry.Id))?" + '$select=Id,FileDirRef,AuthorId,Created')
  if ($moved.FileDirRef -ne $targetFolder -or $moved.AuthorId -ne $entry.AuthorId -or $moved.Created -ne $entry.Created) { throw "Verify journal item $($entry.Id): move did not preserve its expected folder/author/time." }
}
if ($Production) { foreach ($entry in $entries | Where-Object { $_.FSObjType -ne 1 }) { Invoke-QStarRest "Post" "$progressPath/items($($entry.Id))/resetroleinheritance()" | Out-Null } }
Write-Host "Journal ready: $($issues.Count) issue folders, $($moves.Count) entries moved; production ACL reconciliation: $Production."
