# Sideload into Windows PowerPoint by trusting a local network-share catalog.
# Windows desktop Office sideloads from a shared folder registered as a trusted
# catalog, not from a wef folder. Run this once, then add the folder in
#   PowerPoint ▸ File ▸ Options ▸ Trust Center ▸ Trusted Add-in Catalogs
# (or it is added here via the registry), tick "Show in Menu", restart
# PowerPoint, and pick the add-in under  Home ▸ Add-ins ▸ Shared Folder.
#
#   scripts\sideload-windows.ps1            # DEV: points at https://localhost:3000
#   scripts\sideload-windows.ps1 -Hosted    # the production manifest, verbatim
#
# manifest.xml in the repo is the PRODUCTION manifest (see its header); the dev
# copy is derived here by rewriting the hosted origin to the vite dev server.
# For the dev copy the dev server (npm run dev) must be running and its cert
# trusted (npm run certs). Both copies share the add-in <Id>: load one, not both.
#
# UNTESTED on a real Windows PowerPoint as of 2026-10-20 — no machine here has desktop PowerPoint for Windows.
param(
  [string]$ShareDir = "$env:USERPROFILE\.office-addins\pptx-exporter",
  [switch]$Hosted
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $PSScriptRoot
$hostedOrigin = 'https://pptx-exporter-addin.stoatworks-labs.com'
$devOrigin = 'https://localhost:3000'

New-Item -ItemType Directory -Force -Path $ShareDir | Out-Null
$src = Get-Content -Raw (Join-Path $here 'manifest.xml')
if (-not $Hosted) {
  if ($src -notmatch [regex]::Escape($hostedOrigin)) {
    throw "manifest.xml does not mention $hostedOrigin - is it still the production manifest?"
  }
  $src = $src.Replace($hostedOrigin, $devOrigin)
}
Set-Content -Path (Join-Path $ShareDir 'manifest.xml') -Value $src -Encoding UTF8 -NoNewline

# Register the folder as a trusted catalog (per-user, no admin).
$guid = [guid]::NewGuid().ToString()
$base = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
New-Item -Path (Join-Path $base $guid) -Force | Out-Null
$key = Join-Path $base $guid
New-ItemProperty -Path $key -Name 'Id'      -Value $guid            -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Url'     -Value $ShareDir        -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Flags'   -Value 1               -PropertyType DWord  -Force | Out-Null  # 1 = show in menu

$which = if ($Hosted) { "PRODUCTION ($hostedOrigin)" } else { "DEV ($devOrigin)" }
Write-Host "Sideloaded $which manifest -> $ShareDir"
Write-Host "Registered trusted catalog $guid. Restart PowerPoint, then:"
Write-Host "  Home > Add-ins > Shared Folder > Advanced Exporter"
