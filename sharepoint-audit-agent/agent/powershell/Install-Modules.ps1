Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
$mods = @('PnP.PowerShell','ImportExcel')
foreach ($m in $mods) {
  if (-not (Get-Module -ListAvailable -Name $m)) {
    Write-Host "Installing $m ..."
    Install-Module $m -Scope CurrentUser -Force -AllowClobber
  } else {
    Write-Host "$m already present"
  }
}
