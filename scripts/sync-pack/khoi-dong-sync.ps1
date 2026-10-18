# Khoi dong Syncthing va mo thu muc sync (khong cai lai)
$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir 'config.ps1')
. (Join-Path $ScriptDir 'syncthing-setup.ps1')

if (-not (Test-SyncthingInstalled -Config $PackConfig)) {
    Write-Host ''
    Write-Host '[LOI] Syncthing chua duoc cai.' -ForegroundColor Red
    Write-Host 'Chay Cai-Dat-Sync.cmd truoc.' -ForegroundColor Yellow
    Write-Host ''
    exit 1
}

Write-Host ''
Write-Host 'Dang khoi dong Syncthing...' -ForegroundColor Cyan
Start-SyncthingProcess -Config $PackConfig
Start-Sleep -Seconds 2

Write-Host 'Mo thu muc sync...' -ForegroundColor Green
Start-Process $PackConfig.SyncFolder

$url = Get-SyncthingGuiUrl
Write-Host "Trang quan ly: $url" -ForegroundColor Green
try {
    Start-Process $url | Out-Null
} catch {
    Write-Host 'Mo trinh duyet thu cong neu can.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Xong. File trong thu muc Sync se tu dong dong bo.' -ForegroundColor Green
Write-Host ''
