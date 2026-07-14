<#
=====================================================================
 Q-Star Issue Manager — beta provisioning (PnP PowerShell)
=====================================================================
 Creates only the Q-Star lists, fields, indexes, and native Person
 columns. It deliberately does not create Q-Star groups or assign site
 permissions. The web part must have Beta access mode enabled.

 PREREQUISITES
   Install-Module PnP.PowerShell -Scope CurrentUser

 RUN
   ./provision-qstar-beta.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/QStarBeta" -ClientId "<app-guid>"
=====================================================================
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$SiteUrl,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [string]$IssuesList   = "Q-Star Issues",
  [string]$ProgressList = "Q-Star Progress Log",
  [string]$ConfigList   = "Q-Star Config",
  [switch]$PersonAsText
)

$parameters = @{
  SiteUrl = $SiteUrl
  ClientId = $ClientId
  IssuesList = $IssuesList
  ProgressList = $ProgressList
  ConfigList = $ConfigList
  SkipRoleGroups = $true
}

if ($PersonAsText) {
  $parameters.PersonAsText = $true
}

& "$PSScriptRoot/provision-qstar.ps1" @parameters
