param(
  [string]$ManifestPath = "manifests/manifest.staging.xml",
  [string]$SharePath = "$env:USERPROFILE\OfficeAddinCatalog"
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$resolvedManifest = Resolve-Path (Join-Path $projectRoot $ManifestPath)

if (-not (Test-Path $SharePath)) {
  throw "Share path does not exist: $SharePath. Run setup-local-share.ps1 first."
}

$destination = Join-Path $SharePath 'manifest.staging.xml'
Copy-Item -Path $resolvedManifest -Destination $destination -Force

Write-Host "Published manifest to share catalog:"
Write-Host "  $destination"
Write-Host ""
Write-Host "In Excel: Home > Add-ins > More Add-ins > Shared Folder > add manifest.staging.xml"
