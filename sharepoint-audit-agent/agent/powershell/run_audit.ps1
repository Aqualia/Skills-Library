param(
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Tenant,
  [Parameter(Mandatory=$true)][string]$ClientId,
  [Parameter(Mandatory=$true)][string]$CertificatePath,
  [Parameter(Mandatory=$true)][securestring]$CertificatePassword,

  [Parameter(Mandatory=$true)][string]$OriginalScriptPath,  # path to your big script
  [string[]]$InternalDomains,
  [string]$EmitJsonPath,          # where we write audit.json
  [string]$HtmlSearchDir,         # where the big script writes HTML
  [switch]$AutoConfirm,
  [int]$MaxItemsToScan = 50000,
  [int]$BatchSize = 200
)

$ErrorActionPreference = 'Stop'

# 1) Cert-based app-only auth
try {
  Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId `
    -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword
} catch {
  Write-Error "Failed to authenticate with certificate auth. $_"; throw
}

# 2) Prepare args for original script (non-interactive if supported)
if (-not (Test-Path $OriginalScriptPath)) {
  throw "Original script not found at $OriginalScriptPath"
}

$origParamMetadata = $null
$origMetadataAvailable = $false
try {
  $origParamMetadata = (Get-Command -Name $OriginalScriptPath -ErrorAction Stop).Parameters
  $origMetadataAvailable = $true
} catch {
  Write-Verbose "[Wrapper] Could not inspect parameters for $OriginalScriptPath: $_"
}

function ShouldPassParam([string]$name, [bool]$default = $true) {
  if ($origMetadataAvailable) {
    return $origParamMetadata.ContainsKey($name)
  }
  return $default
}

if ($origMetadataAvailable) {
  foreach ($req in @('Url','Tenant','ClientId','CertificatePath','CertificatePassword')) {
    if (-not (ShouldPassParam $req)) {
      throw "Original script at $OriginalScriptPath does not declare required parameter -$req."
    }
  }
}

$origArgs = @()
if (ShouldPassParam 'Tenant') { $origArgs += @('-Tenant', $Tenant) }
if (ShouldPassParam 'ClientId') { $origArgs += @('-ClientId', $ClientId) }
if (ShouldPassParam 'CertificatePath') { $origArgs += @('-CertificatePath', $CertificatePath) }
if (ShouldPassParam 'CertificatePassword') { $origArgs += @('-CertificatePassword', $CertificatePassword) }

if ($AutoConfirm) {
  if (ShouldPassParam 'AutoConfirm') { $origArgs += '-AutoConfirm' }
  elseif ($origMetadataAvailable) { Write-Verbose "[Wrapper] Original script does not accept -AutoConfirm; skipping." }
}
if ($MaxItemsToScan) {
  if (ShouldPassParam 'MaxItemsToScan') { $origArgs += @('-MaxItemsToScan', $MaxItemsToScan) }
  elseif ($origMetadataAvailable) { Write-Verbose "[Wrapper] Original script does not accept -MaxItemsToScan; skipping." }
}
if ($BatchSize) {
  if (ShouldPassParam 'BatchSize') { $origArgs += @('-BatchSize', $BatchSize) }
  elseif ($origMetadataAvailable) { Write-Verbose "[Wrapper] Original script does not accept -BatchSize; skipping." }
}

if (ShouldPassParam 'SiteUrl' $false) {
  $origArgs += @('-SiteUrl', $Url)
}
if (ShouldPassParam 'Url') {
  $origArgs += @('-Url', $Url)
} elseif (-not (ShouldPassParam 'SiteUrl' $false)) {
  # If neither Url nor SiteUrl is declared (unexpected), still attempt Url.
  $origArgs += @('-Url', $Url)
}

if ($InternalDomains -and $InternalDomains.Count -gt 0) {
  if (ShouldPassParam 'InternalDomains') { $origArgs += @('-InternalDomains', $InternalDomains) }
  elseif ($origMetadataAvailable) { Write-Verbose "[Wrapper] Original script does not accept -InternalDomains; skipping." }
}
if ($EmitJsonPath) {
  if (ShouldPassParam 'EmitJsonPath') { $origArgs += @('-EmitJsonPath', $EmitJsonPath) }
  elseif ($origMetadataAvailable) { Write-Verbose "[Wrapper] Original script does not accept -EmitJsonPath; skipping." }
}

Write-Host "[Wrapper] Running original script: $OriginalScriptPath $($origArgs -join ' ')"
$exitCode = 0
try {
  & $OriginalScriptPath @origArgs
  $exitCode = $LASTEXITCODE
} catch {
  Write-Warning "[Wrapper] Original script threw: $_"
  $exitCode = 1
}
if ($exitCode -ne 0) {
  Write-Warning "[Wrapper] Original script exit code: $exitCode (continuing to JSON extraction if possible)"
}

# 3) If JSON already exists, we are done
if ($EmitJsonPath -and (Test-Path $EmitJsonPath)) {
  Write-Host "[Wrapper] Found existing JSON at $EmitJsonPath"
  Disconnect-PnPOnline -ErrorAction SilentlyContinue
  exit 0
}

# 4) Otherwise, search for HTML and extract embedded JSON
if (-not $EmitJsonPath) {
  Write-Host "[Wrapper] No EmitJsonPath specified; skipping JSON extraction."
  Disconnect-PnPOnline -ErrorAction SilentlyContinue
  exit 0
}

if (-not $HtmlSearchDir) { $HtmlSearchDir = (Get-Location).Path }

$htmlFiles = Get-ChildItem -Path $HtmlSearchDir -Recurse -Filter *.htm* -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
if (-not $htmlFiles -or $htmlFiles.Count -eq 0) {
  Write-Warning "[Wrapper] No HTML files found under $HtmlSearchDir; cannot extract JSON."
  Disconnect-PnPOnline -ErrorAction SilentlyContinue
  exit 0
}

$latestHtml = $htmlFiles[0].FullName
Write-Host "[Wrapper] Attempting JSON extraction from: $latestHtml"

$content = Get-Content -Raw -Path $latestHtml -Encoding UTF8
$pattern = '<script[^>]*id=["'']report-data["''][^>]*type=["'']application/json["''][^>]*>(.*?)</script>'
$match = [System.Text.RegularExpressions.Regex]::Match($content, $pattern, 'Singleline, IgnoreCase')
if ($match.Success) {
  $json = $match.Groups[1].Value.Trim()
  try {
    $outDir = Split-Path -Path $EmitJsonPath -Parent
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    $json | Set-Content -Path $EmitJsonPath -Encoding UTF8
    Write-Host "[Wrapper] Wrote $EmitJsonPath"
  } catch {
    Write-Warning "[Wrapper] Failed to write EmitJsonPath: $_"
  }
} else {
  Write-Warning "[Wrapper] Could not find embedded JSON in HTML (id='report-data')."
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
