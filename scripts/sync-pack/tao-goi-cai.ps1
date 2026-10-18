# Tao file zip phan phoi cho nguoi dung (chay tu thu muc sync-pack)
$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutZip = Join-Path $env:USERPROFILE 'Documents\Sync\GoPaperless-Sync-Pack.zip'

$userFiles = @(
    'README.txt', 'Cai-Dat-Sync.cmd', 'Huong-Dan.cmd', 'Khoi-Dong-Sync.cmd', 'HUONG-DAN.html',
    'install.ps1', 'config.ps1', 'syncthing-setup.ps1', 'khoi-dong-sync.ps1', 'legacy'
)
$OutDir = Split-Path $OutZip -Parent

$staging = Join-Path $env:TEMP "GoPaperless-Sync-Pack-$(Get-Date -Format 'yyyyMMddHHmmss')"
New-Item -ItemType Directory -Force -Path $staging | Out-Null

foreach ($name in $userFiles) {
    $src = Join-Path $ScriptDir $name
    if (Test-Path $src) {
        Copy-Item -Recurse -Force $src (Join-Path $staging $name)
    }
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
if (Test-Path $OutZip) { Remove-Item -Force $OutZip }
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $OutZip -Force
Remove-Item -Recurse -Force $staging

Write-Host "Da tao: $OutZip" -ForegroundColor Green
Get-Item $OutZip | Select-Object FullName, Length, LastWriteTime
