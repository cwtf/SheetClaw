#Requires -Version 5.1
<#
.SYNOPSIS
    Installs the SheetClaw Excel add-in for the current user.

.DESCRIPTION
    Downloads the SheetClaw manifest from GitHub Pages and registers it as a
    trusted add-in catalog in Excel. After running this script, open Excel and
    enable SheetClaw once via Insert > Add-ins > My Add-ins > Shared Folder.
#>

$ErrorActionPreference = 'Stop'

$ManifestUrl  = 'https://cwtf.github.io/SheetClaw/manifest.xml'
$InstallDir   = Join-Path $env:LOCALAPPDATA 'SheetClaw'
$ManifestPath = Join-Path $InstallDir 'manifest.xml'
$CatalogKey   = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\SheetClaw'

function Write-Step([string]$msg) {
    Write-Host "  $msg" -ForegroundColor Cyan
}

Write-Host ''
Write-Host 'SheetClaw Installer' -ForegroundColor White
Write-Host '-------------------'
Write-Host ''

# Warn if Excel is open — registry changes take effect only after restart
$excelProc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Warning 'Excel is currently open. Close it before continuing, then press Enter.'
    $null = Read-Host
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        Write-Error 'Excel is still running. Please close it and re-run this script.'
        exit 1
    }
}

# Create install directory
Write-Step "Creating install directory: $InstallDir"
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

# Download manifest
Write-Step 'Downloading manifest...'
try {
    Invoke-WebRequest -Uri $ManifestUrl -OutFile $ManifestPath -UseBasicParsing
} catch {
    Write-Error "Failed to download manifest from $ManifestUrl`n$_"
    exit 1
}

# Validate it looks like XML
$content = Get-Content $ManifestPath -Raw -ErrorAction SilentlyContinue
if ($content -notmatch '<OfficeApp') {
    Write-Error "Downloaded file does not look like a valid Office manifest: $ManifestPath"
    exit 1
}
Write-Step "Manifest saved to: $ManifestPath"

# Register as a trusted catalog
Write-Step 'Registering trusted catalog in Excel...'
if (-not (Test-Path $CatalogKey)) {
    New-Item -Path $CatalogKey -Force | Out-Null
}
Set-ItemProperty -Path $CatalogKey -Name 'Url'   -Value $InstallDir
Set-ItemProperty -Path $CatalogKey -Name 'Flags' -Value 1 -Type DWord

Write-Host ''
Write-Host 'Done!' -ForegroundColor Green
Write-Host ''
Write-Host 'To activate SheetClaw:'
Write-Host '  1. Open Excel'
Write-Host '  2. Insert > Add-ins > My Add-ins'
Write-Host '  3. Click the "Shared Folder" tab'
Write-Host '  4. Select SheetClaw and click Add'
Write-Host ''
Write-Host 'SheetClaw will appear automatically in all future Excel sessions.'
Write-Host ''
