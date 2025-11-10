<#
.SYNOPSIS
  SharePoint Audit Script (MVP) — certificate auth, Sites.Selected compatible.
  Scans a target site for key permission risks and emits:
   - HTML report (if your existing HTML export logic is present),
   - Standalone JSON (EmitJsonPath) for downstream analysis.

.DESCRIPTION
  This script connects app-only using a certificate, walks webs/lists/items
  up to MaxItemsToScan, computes high-level metrics, and builds a $reportObj
  suitable for HTML/JSON export. Internal/external classification is driven
  by -InternalDomains at runtime (no tenant-specific hardcoding).

.PARAMETER Url
  Full site URL to audit (e.g., https://contoso.sharepoint.com/sites/Finance)

.PARAMETER Tenant
  Tenant GUID or tenant domain (contoso.onmicrosoft.com)

.PARAMETER ClientId
  Application (client) ID for your Entra app

.PARAMETER CertificatePath
  Path to PFX certificate file

.PARAMETER CertificatePassword
  SecureString password for the PFX

.PARAMETER InternalDomains
  Array of email domains considered "internal" (e.g., aqualia.ie, aqualia.onmicrosoft.com)

.PARAMETER EmitJsonPath
  Path to write a standalone JSON representation of the report object

.PARAMETER AutoConfirm
  Suppresses prompts during scan (non-interactive mode)

.PARAMETER MaxItemsToScan
  Upper bound for list item scanning across the site

.PARAMETER BatchSize
  Page size for list item enumeration
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Tenant,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$CertificatePath,
  [Parameter(Mandatory=$true)][SecureString]$CertificatePassword,
  [string[]]$InternalDomains,
  [string]$EmitJsonPath,
  [switch]$AutoConfirm,
  [int]$MaxItemsToScan = 50000,
  [int]$BatchSize = 200
)

# ---------- Helpers ----------
function Test-IsInternal([string]$Email, [string[]]$Domains) {
  if (-not $Email) { return $false }
  if (-not $Domains -or $Domains.Count -eq 0) { return $false }
  $lower = $Email.ToLower()
  foreach ($d in $Domains) {
    $suffix = "@" + ($d.TrimStart('@').ToLower())
    if ($lower.EndsWith($suffix)) { return $true }
  }
  return $false
}

function Invoke-ForEachListItemStreaming([Microsoft.SharePoint.Client.List]$List, [int]$PageSize, [scriptblock]$Handler) {
  $position = $null
  do {
    $items = Get-PnPListItem -List $List -PageSize $PageSize -ListItemCollectionPosition $position -ScriptBlock { param($items) $items }
    $position = $items.ListItemCollectionPosition
    foreach ($it in $items) { & $Handler $it }
  } while ($position -ne $null)
}

# ---------- Connect (App-only with cert) ----------
try {
  Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId `
    -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword
}
catch {
  Write-Error "Failed to authenticate with certificate auth. $_"
  throw
}

$site = Get-PnPSite
$web  = Get-PnPWeb
Write-Host "Connected to $($web.Url)"

# ---------- Metrics scaffold ----------
$metrics = [ordered]@{
  siteUrl                    = $Url
  scannedAt                  = (Get-Date).ToString("s")
  itemsWithUniquePermissions = 0
  externalUsers              = 0
  webDirectAssignments       = 0
  orphanedGroups             = 0
  anyoneOrEveryoneAtWeb      = $false
  externalOwnerPresent       = $false
  totalLists                 = 0
  totalItemsScanned          = 0
}

$findings = New-Object System.Collections.ArrayList
$details  = New-Object System.Collections.ArrayList

# ---------- Web-scope assignments ----------
try {
  $roleAssignments = Get-PnPRoleAssignment -Scope Web -ErrorAction SilentlyContinue
  foreach ($ra in $roleAssignments) {
    if ($ra.MemberTitle -match "Everyone" -or $ra.MemberTitle -match "Anyone") {
      $metrics.anyoneOrEveryoneAtWeb = $true
      [void]$findings.Add(@{ level = "Critical"; message = "Anyone/Everyone present at Web scope" })
    }
    # Count direct user assignments (not groups)
    if ($ra.PrincipalType -eq "User") {
      $metrics.webDirectAssignments++
    }
  }
} catch {
  Write-Warning "Failed to read web role assignments: $_"
}

# ---------- SharePoint groups sanity checks ----------
try {
  $spGroups = Get-PnPGroup -ErrorAction SilentlyContinue
  foreach ($g in $spGroups) {
    if (-not $g.OwnerTitle -or $g.OwnerTitle.Trim().Length -eq 0) {
      $metrics.orphanedGroups++
    }
  }
} catch {
  Write-Warning "Failed to enumerate SP groups: $_"
}

# ---------- Enumerate lists and items ----------
try {
  $lists = Get-PnPList -Includes RootFolder -ErrorAction SilentlyContinue | Where-Object { -not $_.Hidden }
  $metrics.totalLists = ($lists | Measure-Object).Count
} catch {
  Write-Warning "Failed to enumerate lists: $_"
  $lists = @()
}

$script:scannedCount = 0
foreach ($list in $lists) {
  if ($script:scannedCount -ge $MaxItemsToScan) { break }

  $listName = $list.Title
  $listUrl  = $list.RootFolder.ServerRelativeUrl

  # Stream items
  try {
    Invoke-ForEachListItemStreaming -List $list -PageSize $BatchSize -Handler {
      param($item)

      if ($script:scannedCount -ge $MaxItemsToScan) { return }
      $script:scannedCount++

      # Unique permissions?
      $hasUnique = $false
      try {
        $hasUnique = (Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments).HasUniqueRoleAssignments
      } catch { $hasUnique = $false }

      if ($hasUnique) {
        $metrics.itemsWithUniquePermissions++

        # Try to look at role assignments at item level
        try {
          $assigns = Get-PnPProperty -ClientObject $item -Property RoleAssignments
          foreach ($ra in $assigns) {
            $m = Get-PnPProperty -ClientObject $ra -Property Member
            $title = $m.Title
            $type  = $m.PrincipalType.ToString()
            $isExternal = $false

            if ($type -eq "User" -and $m.Email) {
              $isExternal = -not (Test-IsInternal -Email $m.Email -Domains $InternalDomains)
              if ($isExternal) { $metrics.externalUsers++ }
            }

            # External owner?
            $binds = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings
            foreach ($b in $binds) {
              if ($b.Name -match "Full Control" -and $isExternal) {
                $metrics.externalOwnerPresent = $true
                [void]$findings.Add(@{ level = "Critical"; message = "External identity with Full Control on item"; path = $listUrl })
              }
            }
          }
        } catch {
          Write-Verbose "Failed to inspect item role assignments for $listName: $_"
        }
      }

      # Optionally, collect detail rows (keep light in MVP)
      $null = $details.Add(@{
        list   = $listName
        url    = $listUrl
        unique = $hasUnique
      })
    }
  } catch {
    Write-Warning "Failed to stream list '$listName': $_"
  }
}

$metrics.totalItemsScanned = $script:scannedCount

# ---------- Compose report object ----------
$reportObj = [ordered]@{
  version = "mvp-1"
  site    = $Url
  metrics = $metrics
  notes   = @(
    if (-not $InternalDomains -or $InternalDomains.Count -eq 0) { "Internal domains not provided; internal/external classification limited." }
  )
  details = $details
  findings = $findings
}

# ---------- HTML export (placeholder-friendly) ----------
# If you already generate an HTML report elsewhere in your original script,
# keep that logic. This MVP does not enforce an HTML format, but $reportObj
# is suitable for rendering.

# ---------- JSON export ----------
if ($EmitJsonPath) {
  try {
    ($reportObj | ConvertTo-Json -Depth 20) | Set-Content -Path $EmitJsonPath -Encoding UTF8
    Write-Host "[EmitJson] Wrote $EmitJsonPath"
  } catch {
    Write-Warning "Failed to write EmitJsonPath: $_"
  }
}

Write-Host "Audit complete. Items scanned: $($metrics.totalItemsScanned). Unique perms: $($metrics.itemsWithUniquePermissions). External users: $($metrics.externalUsers)."
