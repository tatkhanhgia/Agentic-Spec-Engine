# GoPaperless Sync Pack - one-click installer (Syncthing)

param(
    [ValidateSet('Syncthing', 'Ssh', 'Menu')]
    [string]$Mode = 'Syncthing',
    [string[]]$RemoteDeviceIds = @(),
    [switch]$SkipPrompt
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'config.ps1')
. (Join-Path $ScriptDir 'syncthing-setup.ps1')

function Show-Banner {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Green
    Write-Host '   GO PAPERLESS - CAI DAT DONG BO FILE' -ForegroundColor Green
    Write-Host '========================================' -ForegroundColor Green
    Write-Host ''
    Write-Host 'Dong bo tu dong giua nhieu may (qua Tailscale)'
    Write-Host "Thu muc sync: $($PackConfig.SyncFolder)"
    Write-Host ''
}

function Show-SyncthingResult {
    param($Result)

    Write-Host ''
    Write-Host '========================================' -ForegroundColor Green
    Write-Host '   CAI DAT XONG!' -ForegroundColor Green
    Write-Host '========================================' -ForegroundColor Green
    Write-Host ''
    Write-Ok "Thu muc sync: $($Result.SyncFolder)"
    Write-Ok "Trang quan ly: $($Result.GuiUrl)"
    Write-Host ''
    Write-Host 'DEVICE ID MAY NAY (copy gui cho may khac):' -ForegroundColor Yellow
    Write-Host $Result.DeviceId -ForegroundColor White
    Write-Host ''
    Write-Host 'Tren may khac: chay Cai-Dat-Sync.cmd va dan Device ID o tren.' -ForegroundColor Cyan
    Write-Host 'Can huong dan chi tiet? Chay Huong-Dan.cmd' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Sau do tha file vao thu muc Sync la tu dong dong bo.' -ForegroundColor Green
    Write-Host ''

    try {
        Start-Process $Result.GuiUrl | Out-Null
    } catch {
        Write-Warn "Mo trinh duyet thu cong: $($Result.GuiUrl)"
    }
}

function Read-DeviceIds {
    $ids = @()
    foreach ($id in (Get-PeerDeviceIds)) {
        $ids += $id
    }

    if ($ids.Count -gt 0) {
        Write-Host 'Device ID da luu san trong config:' -ForegroundColor Cyan
        foreach ($id in $ids) {
            Write-Host "  $id"
        }
        Write-Host ''
    }

    Write-Host 'Nhap Device ID cua may khac (Enter = xong, bo qua):' -ForegroundColor Yellow
    while ($true) {
        $input = Read-Host 'Device ID'
        if (-not $input) { break }
        $trimmed = $input.Trim()
        if ($ids -contains $trimmed) {
            Write-Warn 'Device ID nay da co, bo qua.'
            continue
        }
        $ids += $trimmed
    }

    return $ids
}

function Install-SshMode {
    $legacyDir = Join-Path $ScriptDir 'legacy'
    Write-Info 'Che do nang cao: SSH + Unison (legacy)'
    $installUnison = Join-Path $legacyDir 'install-unison.cmd'
    $setupSync = Join-Path $legacyDir 'setup-sync.cmd'

    if (-not (Test-Path $installUnison)) {
        throw "Khong tim thay: $installUnison"
    }

    Write-Host ''
    Write-Warn 'Che do nay can WSL + Unison. Chi ho tro 2 may.'
    Write-Host ''

    cmd /c "`"$installUnison`""
    cmd /c "`"$setupSync`""

    Write-Host ''
    Write-Ok 'SSH sync da setup. Chay sync bang: legacy\sync-folder.cmd'
    Write-Host ''
}

Show-Banner

if ($Mode -eq 'Menu') {
    Write-Host 'Chon che do cai dat:' -ForegroundColor Cyan
    Write-Host '  [1] Syncthing - mac dinh (tu dong sync nhieu may)'
    Write-Host '  [2] SSH/Unison - nang cao (legacy, chi 2 may)'
    Write-Host ''
    $choice = Read-Host 'Nhap 1 hoac 2 (mac dinh: 1)'
    if (-not $choice -or $choice -eq '1') {
        $Mode = 'Syncthing'
    } else {
        $Mode = 'Ssh'
    }
}

if ($Mode -eq 'Syncthing') {
    if ($RemoteDeviceIds.Count -eq 0) {
        if ($SkipPrompt) {
            $RemoteDeviceIds = @(Get-PeerDeviceIds)
        } else {
            $RemoteDeviceIds = Read-DeviceIds
        }
    }

    $result = Install-SyncthingMode -Config $PackConfig -RemoteDeviceIds $RemoteDeviceIds
    Show-SyncthingResult -Result $result
    exit 0
}

if ($Mode -eq 'Ssh') {
    Install-SshMode
    exit 0
}

throw "Che do khong hop le: $Mode"
