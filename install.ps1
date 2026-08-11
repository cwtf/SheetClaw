#Requires -Version 5.1
<#
.SYNOPSIS
    Installs the SheetClaw Excel add-in for the current user.

.DESCRIPTION
    Places the SheetClaw manifest in a local folder, shares that folder over SMB,
    and registers the share as a trusted add-in catalog in Excel.

    The share is not optional. Excel's trusted catalog accepts a network path
    (\\PC\Share) only - a plain local path like C:\Users\... is accepted by the
    registry but silently yields an empty Shared Folder tab. Creating a share
    needs elevation, so this script relaunches itself as administrator unless
    -SkipShare is passed.

.PARAMETER ManifestUrl
    Where to download the manifest from. Defaults to the GitHub Pages deployment.

.PARAMETER LocalManifest
    Use a manifest already on disk instead of downloading - e.g. .\manifest.xml
    when working from a clone. Skips the network entirely.

.PARAMETER SkipShare
    Register the catalog without creating the SMB share. Only useful if you have
    already shared the folder yourself; otherwise the add-in will not appear.

.EXAMPLE
    .\install.ps1

.EXAMPLE
    .\install.ps1 -LocalManifest .\manifest.xml
#>

[CmdletBinding()]
param(
    [string]$ManifestUrl = 'https://cwtf.github.io/SheetClaw/manifest.xml',
    [string]$LocalManifest,
    [string]$ShareName = 'SheetClaw',
    [switch]$SkipShare,
    [switch]$Relaunched
)

$ErrorActionPreference = 'Stop'

$InstallDir   = Join-Path $env:LOCALAPPDATA 'SheetClaw'
$ManifestPath = Join-Path $InstallDir 'manifest.xml'
$CatalogKey   = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\SheetClaw'
$CatalogId    = '{F58DEFDB-FB64-4DA7-9F25-761DAB45DE9B}'

function Write-Step([string]$msg) {
    Write-Host "  $msg" -ForegroundColor Cyan
}

function Write-Note([string]$msg) {
    Write-Host "  $msg" -ForegroundColor DarkGray
}

# Write-Error under $ErrorActionPreference='Stop' throws before any following
# exit, which buries the message under a PowerShell stack trace. Fail plainly.
function Fail([string]$msg) {
    Write-Host ''
    Write-Host "ERROR: $msg" -ForegroundColor Red
    Write-Host ''
    exit 1
}

function Test-Administrator {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Write-Host ''
Write-Host 'SheetClaw Installer' -ForegroundColor White
Write-Host '-------------------'
Write-Host ''

# ── Excel must be closed: the catalog registration is read at startup ───────
$excelProc = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Warning 'Excel is currently open. Close it, then press Enter to continue.'
    $null = Read-Host
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        Fail 'Excel is still running. Close it and re-run this script.'
    }
}

# ── Elevate for the share, unless told not to ──────────────────────────────
if (-not $SkipShare -and -not (Test-Administrator)) {
    if ($Relaunched) {
        Fail 'Elevation was declined. Re-run as administrator, or use -SkipShare if the folder is already shared.'
    }
    Write-Step 'Creating the network share needs administrator rights - relaunching...'

    $argList = @(
        '-NoProfile'
        '-ExecutionPolicy', 'Bypass'
        '-File', ('"{0}"' -f $PSCommandPath)
        '-ManifestUrl', ('"{0}"' -f $ManifestUrl)
        '-ShareName', ('"{0}"' -f $ShareName)
        '-Relaunched'
    )
    if ($LocalManifest) {
        $argList += @('-LocalManifest', ('"{0}"' -f (Resolve-Path $LocalManifest).Path))
    }

    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -PassThru -Wait
    } catch {
        Fail 'Could not relaunch as administrator. Right-click PowerShell, Run as administrator, and try again.'
    }
    exit $proc.ExitCode
}

# ── Install directory ──────────────────────────────────────────────────────
Write-Step "Install directory: $InstallDir"
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

# ── Obtain the manifest ────────────────────────────────────────────────────
# Staged through a temp file so a bad download never overwrites a working
# manifest, and so the HTML of a 404 page cannot land where Excel reads it.
$tempManifest = Join-Path ([IO.Path]::GetTempPath()) ("sheetclaw-manifest-{0}.xml" -f [guid]::NewGuid())

if ($LocalManifest) {
    if (-not (Test-Path $LocalManifest)) {
        Fail "No manifest at: $LocalManifest"
    }
    Write-Step "Using local manifest: $LocalManifest"
    Copy-Item -Path $LocalManifest -Destination $tempManifest -Force
} else {
    Write-Step "Downloading manifest from $ManifestUrl"
    try {
        # PS 5.1 on older Windows still defaults to TLS 1.0, which GitHub rejects.
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $ManifestUrl -OutFile $tempManifest -UseBasicParsing -TimeoutSec 60
    } catch {
        # Report the status only. The body of a GitHub Pages 404 is a full HTML
        # page, and dumping it buries the actual problem.
        $status = $null
        if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
            try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = $null }
        }
        Remove-Item $tempManifest -Force -ErrorAction SilentlyContinue

        if ($status -eq 404) {
            Fail @"
Manifest not found at $ManifestUrl (HTTP 404).

The GitHub Pages deployment may not have published it yet. Either wait for the
Deploy workflow to finish, or install from your clone instead:

    .\install.ps1 -LocalManifest .\manifest.xml
"@
        }
        if ($status) {
            Fail "Download failed with HTTP $status from $ManifestUrl"
        }
        Fail "Could not reach $ManifestUrl - $($_.Exception.Message)"
    }
}

# ── Validate before it goes anywhere Excel will read ───────────────────────
$content = Get-Content $tempManifest -Raw -ErrorAction SilentlyContinue
if ($content -notmatch '<OfficeApp') {
    Remove-Item $tempManifest -Force -ErrorAction SilentlyContinue
    Fail "The downloaded file is not an Office manifest (no <OfficeApp> element). Check $ManifestUrl in a browser."
}

Move-Item -Path $tempManifest -Destination $ManifestPath -Force
Write-Step "Manifest saved to: $ManifestPath"

# ── Share the folder; Excel needs a UNC path, not a local one ──────────────
$catalogUrl = $null

if ($SkipShare) {
    Write-Note 'Skipping share creation (-SkipShare).'
    $existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if ($existing) {
        $catalogUrl = "\\$env:COMPUTERNAME\$ShareName"
    } else {
        Fail @"
-SkipShare was passed but no share named '$ShareName' exists.

Excel's trusted catalog only accepts a network path. Share $InstallDir manually
(right-click > Properties > Sharing), then re-run with -SkipShare.
"@
    }
} else {
    $existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if ($existing) {
        if ($existing.Path -ne $InstallDir) {
            Write-Step "Repointing existing share '$ShareName' at $InstallDir"
            Remove-SmbShare -Name $ShareName -Force | Out-Null
            New-SmbShare -Name $ShareName -Path $InstallDir -ReadAccess $env:USERNAME | Out-Null
        } else {
            Write-Step "Share '$ShareName' already points at the install directory"
        }
    } else {
        Write-Step "Creating read-only share '$ShareName' for $InstallDir"
        try {
            New-SmbShare -Name $ShareName -Path $InstallDir -ReadAccess $env:USERNAME | Out-Null
        } catch {
            Fail "Could not create the share: $($_.Exception.Message)"
        }
    }
    $catalogUrl = "\\$env:COMPUTERNAME\$ShareName"
}

# ── Register the trusted catalog ───────────────────────────────────────────
Write-Step "Registering trusted catalog: $catalogUrl"
if (-not (Test-Path $CatalogKey)) {
    New-Item -Path $CatalogKey -Force | Out-Null
}
Set-ItemProperty -Path $CatalogKey -Name 'Id'    -Value $CatalogId
Set-ItemProperty -Path $CatalogKey -Name 'Url'   -Value $catalogUrl
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
if ($Relaunched) {
    Write-Host 'Press Enter to close...'
    $null = Read-Host
}
