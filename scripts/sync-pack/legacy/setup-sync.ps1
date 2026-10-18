# One-time setup for bidirectional SSH sync (Windows + WSL + Unison).
# Run on BOTH machines, or run with -RemoteSetup from the machine you SSH from.
#
# Usage:
#   .\setup-sync.ps1              # local setup only
#   .\setup-sync.ps1 -RemoteSetup # also try to configure the other machine via SSH

param(
    [switch]$RemoteSetup
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'config.ps1')

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Test-WslAvailable {
    try {
        wsl --status | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Install-UnisonInWsl {
    Write-Step 'Installing Unison in WSL (Ubuntu)...'

    $installed = wsl bash -lc "command -v unison >/dev/null 2>&1"
    if ($LASTEXITCODE -eq 0) {
        wsl bash -lc "echo 'Unison already installed:'; unison -version | head -1"
        return
    }

    Write-Host 'Trying non-interactive install...' -ForegroundColor Yellow
    cmd /c "wsl sudo -n apt-get update -qq 2>nul"
    cmd /c "wsl sudo -n apt-get install -y unison openssh-client 2>nul"
    wsl bash -lc "command -v unison >/dev/null 2>&1"
    if ($LASTEXITCODE -eq 0) {
        wsl bash -lc "echo 'Installed:'; unison -version | head -1"
        return
    }

    Write-Warning 'Unison is not installed yet (WSL sudo password required).'
    Write-Host @"

Run this once in Ubuntu (WSL), then re-run setup-sync.ps1:

  wsl
  sudo apt-get update
  sudo apt-get install -y unison openssh-client
  exit

Or run: .\install-unison.ps1
"@
}

function Install-UnisonProfile {
    Write-Step 'Installing Unison profile in WSL...'
    $profileSource = Join-Path $ScriptDir 'unison\sync.prf'
    if (-not (Test-Path $profileSource)) {
        throw "Missing profile: $profileSource"
    }

    $wslSource = ConvertTo-WslPath $profileSource
    wsl bash -lc "mkdir -p ~/.unison && cp '$wslSource' ~/.unison/sync.prf && echo 'Profile installed: ~/.unison/sync.prf'"
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to install Unison profile in WSL'
    }
}

function Add-SshHostKey {
    Write-Step 'Adding SSH host key for remote machine...'
    $knownHosts = Join-Path $env:USERPROFILE '.ssh\known_hosts'
    $sshDir = Split-Path $knownHosts
    if (-not (Test-Path $sshDir)) {
        New-Item -ItemType Directory -Force -Path $sshDir | Out-Null
    }

    $scanOutput = cmd /c "ssh-keyscan -H $($SyncConfig.RemoteHost) 2>nul"
    $hostKey = $scanOutput | Where-Object { $_ -and $_ -notmatch '^#' }
    if (-not $hostKey) {
        Write-Warning "Could not fetch host key for $($SyncConfig.RemoteHost)."
        Write-Host "Connect manually once: ssh $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)"
        return
    }

    $existing = if (Test-Path $knownHosts) { Get-Content $knownHosts -Raw } else { '' }
    if ($existing -notmatch [regex]::Escape($SyncConfig.RemoteHost)) {
        Add-Content -Path $knownHosts -Value ($hostKey -join "`n")
        Write-Host "Host key saved to $knownHosts"
    } else {
        Write-Host 'Host key already present in known_hosts'
    }

    wsl bash -lc "mkdir -p ~/.ssh && chmod 700 ~/.ssh"
    wsl bash -lc "ssh-keyscan -H $($SyncConfig.RemoteHost) 2>/dev/null >> ~/.ssh/known_hosts || true"
}

function Ensure-LocalSyncFolder {
    Write-Step 'Creating local Sync folder...'
    New-Item -ItemType Directory -Force -Path $SyncConfig.SyncFolder | Out-Null
    $readme = Join-Path $SyncConfig.SyncFolder 'README.txt'
    if (-not (Test-Path $readme)) {
        @"
Sync folder
===========
Files placed here sync bidirectionally with the other Windows machine.

Remote: $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)
Path:   $($SyncConfig.SyncFolder)

Run sync once:  scripts\sync-pack\sync-folder.ps1
Watch mode:     scripts\sync-pack\sync-folder.ps1 -Watch
"@ | Set-Content -Path $readme -Encoding UTF8
    }
    Write-Host "Local folder ready: $($SyncConfig.SyncFolder)"
}

function Invoke-RemoteSetup {
    Write-Step 'Setting up remote machine via SSH...'
    $remote = "$($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)"
    $syncPath = $SyncConfig.SyncFolder
    $remoteSetupScript = Join-Path $ScriptDir 'setup-sync.ps1'

    ssh $remote "powershell -NoProfile -Command `"New-Item -ItemType Directory -Force -Path '$syncPath' | Out-Null; Write-Host 'Remote Sync folder OK'`""
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Could not create remote folder via SSH."
        return
    }

    Write-Host "Running remote setup script..."
    ssh $remote "powershell -NoProfile -ExecutionPolicy Bypass -File `"$remoteSetupScript`""
    if ($LASTEXITCODE -ne 0) {
        Write-Warning 'Remote setup did not complete automatically.'
        Write-Host @"

On the REMOTE machine ($remote), run:

  cd C:\Users\Admin\Documents\NetBeansProjects\MyProject\api-document-specification\scripts\sync-pack
  .\setup-sync.ps1

If this repo is not on the remote yet, copy the scripts folder there first.
"@
    }
}

Write-Host 'GoPaperless SSH Sync Setup' -ForegroundColor Green
Write-Host "Remote: $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)"
Write-Host "Folder: $($SyncConfig.SyncFolder)"

if (-not (Test-WslAvailable)) {
    throw @"
WSL is required but not available.

Install WSL first:
  wsl --install
Then reboot and run this script again.
"@
}

Ensure-LocalSyncFolder
Install-UnisonInWsl
Install-UnisonProfile
Add-SshHostKey

if ($RemoteSetup) {
    Invoke-RemoteSetup
}

Write-Step 'Setup complete'
Write-Host @"

Next steps:
  1. Run the same setup on the other machine:
       .\setup-sync.ps1
  2. Test SSH:
       ssh $($SyncConfig.RemoteUser)@$($SyncConfig.RemoteHost)
  3. Sync once:
       .\sync-folder.ps1
  4. Continuous sync:
       .\sync-folder.ps1 -Watch

Sync folder: $($SyncConfig.SyncFolder)
"@
