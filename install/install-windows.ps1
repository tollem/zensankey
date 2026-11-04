<#
.SYNOPSIS
  Installs (or removes) the ZenSankey add-ins for Excel and PowerPoint on Windows.

.DESCRIPTION
  Copies ZenSankey.xlam and ZenSankey.ppam to the user's Office add-in folder
  (%APPDATA%\Microsoft\AddIns), removes the "downloaded from the Internet" mark and
  registers both add-ins so that they load automatically:
    - Excel:      HKCU\Software\Microsoft\Office\<version>\Excel\Options  (OPEN values)
    - PowerPoint: HKCU\Software\Microsoft\Office\<version>\PowerPoint\AddIns\ZenSankey
  Close Excel and PowerPoint before running this script. No administrator rights are needed.

.PARAMETER Uninstall
  Removes the registrations and the copied files.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install\install-windows.ps1
  powershell -ExecutionPolicy Bypass -File install\install-windows.ps1 -Uninstall
#>
param([switch]$Uninstall)

$ErrorActionPreference = "Stop"
$name = "ZenSankey"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path (Split-Path -Parent $here) "dist"
if (-not (Test-Path (Join-Path $distDir "$name.xlam"))) { $distDir = $here }
$xlamSrc = Join-Path $distDir "$name.xlam"
$ppamSrc = Join-Path $distDir "$name.ppam"
$addinDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$xlam = Join-Path $addinDir "$name.xlam"
$ppam = Join-Path $addinDir "$name.ppam"

function Get-OfficeVersions {
  $versions = @()
  foreach ($v in "16.0", "15.0", "14.0") {
    if (Test-Path "HKCU:\Software\Microsoft\Office\$v") { $versions += $v }
  }
  if ($versions.Count -eq 0) { $versions = @("16.0") }
  return $versions
}

function Register-Excel($version, $path) {
  $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
  if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
  $props = Get-ItemProperty -Path $key
  foreach ($p in $props.PSObject.Properties) {
    if ($p.Name -match '^OPEN\d*$' -and $p.Value -like "*$name.xlam*") { return }
  }
  $n = ""; $i = 0
  while ($null -ne $props.PSObject.Properties["OPEN$n"]) { $i++; $n = "$i" }
  New-ItemProperty -Path $key -Name "OPEN$n" -Value "/R `"$path`"" -PropertyType String -Force | Out-Null
}

function Unregister-Excel($version) {
  $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
  if (-not (Test-Path $key)) { return }
  $props = Get-ItemProperty -Path $key
  foreach ($p in $props.PSObject.Properties) {
    if ($p.Name -match '^OPEN\d*$' -and $p.Value -like "*$name.xlam*") { Remove-ItemProperty -Path $key -Name $p.Name }
  }
}

function Register-PowerPoint($version, $path) {
  $key = "HKCU:\Software\Microsoft\Office\$version\PowerPoint\AddIns\$name"
  New-Item -Path $key -Force | Out-Null
  New-ItemProperty -Path $key -Name "Path" -Value $path -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $key -Name "AutoLoad" -Value 1 -PropertyType DWord -Force | Out-Null
}

function Unregister-PowerPoint($version) {
  $key = "HKCU:\Software\Microsoft\Office\$version\PowerPoint\AddIns\$name"
  if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
}

if ($Uninstall) {
  foreach ($v in Get-OfficeVersions) { Unregister-Excel $v; Unregister-PowerPoint $v }
  foreach ($f in $xlam, $ppam) { if (Test-Path $f) { Remove-Item $f -Force } }
  Write-Host "$name removed. Restart Excel and PowerPoint."
  exit 0
}

if (-not (Test-Path $xlamSrc) -or -not (Test-Path $ppamSrc)) {
  Write-Error "Add-in files not found next to this script or in ..\dist ($name.xlam, $name.ppam)."
}
New-Item -ItemType Directory -Path $addinDir -Force | Out-Null
Copy-Item $xlamSrc $xlam -Force
Copy-Item $ppamSrc $ppam -Force
foreach ($f in $xlam, $ppam) { Unblock-File -Path $f -ErrorAction SilentlyContinue }
foreach ($v in Get-OfficeVersions) {
  Register-Excel $v $xlam
  Register-PowerPoint $v $ppam
}
Write-Host "$name installed:"
Write-Host "  $xlam"
Write-Host "  $ppam"
Write-Host "Start Excel or PowerPoint: a 'Sankey' tab appears in the ribbon."
Write-Host "If macros are blocked by policy, add $addinDir to the Trusted Locations (File > Options > Trust Center)."
