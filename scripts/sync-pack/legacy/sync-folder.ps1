# Bidirectional sync for Documents\Sync via Unison over SSH.
#
# Usage:
#   .\sync-folder.ps1           # sync once
#   .\sync-folder.ps1 -Watch    # keep watching and syncing
#   .\sync-folder.ps1 -Interactive  # ask before resolving conflicts

param(
    [switch]$Watch,
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'config.ps1')

function Test-WslAvailable {
    try {
        wsl --status | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Test-UnisonInstalled {
    wsl bash -lc 'command -v unison' 2>$null
    return $LASTEXITCODE -eq 0
}

function Ensure-Ready {
    if (-not (Test-WslAvailable)) {
        throw 'WSL is not available. Run .\setup-sync.ps1 first.'
    }
    if (-not (Test-UnisonInstalled)) {
        throw 'Unison is not installed in WSL. Run .\setup-sync.ps1 first.'
    }
    if (-not (Test-Path $SyncConfig.SyncFolder)) {
        New-Item -ItemType Directory -Force -Path $SyncConfig.SyncFolder | Out-Null
    }

    $profilePath = Join-Path $env:USERPROFILE '.unison\sync.prf'
    wsl bash -lc 'test -f ~/.unison/sync.prf' 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw 'Unison profile missing in WSL. Run .\setup-sync.ps1 first.'
    }
}

function Invoke-Unison {
    param(
        [string[]]$ExtraArgs
    )

    $argsLine = ($ExtraArgs -join ' ')
    Write-Host "Syncing: $($SyncConfig.SyncFolder)" -ForegroundColor Green
    Write-Host "Remote:  $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)" -ForegroundColor Green
    Write-Host ''

    wsl bash -lc "unison sync $argsLine"
    if ($LASTEXITCODE -ne 0) {
        throw "Unison exited with code $LASTEXITCODE"
    }
}

Ensure-Ready

$extra = @()
if (-not $Interactive) {
    $extra += '-auto'
}
if ($Watch) {
    $extra += '-repeat' , 'watch'
}

try {
    Invoke-Unison -ExtraArgs $extra
    Write-Host "`nSync finished." -ForegroundColor Green
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host @"

Tips:
  - First sync? Run: .\setup-sync.ps1
  - SSH failed? Test: ssh $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)
  - Conflicts? Run: .\sync-folder.ps1 -Interactive
"@
    exit 1
}
