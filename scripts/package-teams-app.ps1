# Packages the Teams app manifest + icons into a sideload-ready .zip.
#
# Usage:
#   pwsh ./scripts/package-teams-app.ps1            # dev
#   pwsh ./scripts/package-teams-app.ps1 -Env prod  # prod

[CmdletBinding()]
param(
  [string] $Env = "dev",
  [string] $ManifestPath = (Join-Path $PSScriptRoot "..\appPackage\manifest.json"),
  [string] $OutputDir   = (Join-Path $PSScriptRoot "..\appPackage\build")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$zipPath = Join-Path $OutputDir "appPackage.$Env.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

$manifestDir = Split-Path -Parent $ManifestPath
$colorPng    = Join-Path $manifestDir "color.png"
$outlinePng  = Join-Path $manifestDir "outline.png"

foreach ($p in @($ManifestPath, $colorPng, $outlinePng)) {
  if (-not (Test-Path $p)) {
    throw "Required file not found: $p"
  }
}

Compress-Archive -Path $ManifestPath, $colorPng, $outlinePng `
                 -DestinationPath $zipPath -Force

Write-Host "Wrote $zipPath"
