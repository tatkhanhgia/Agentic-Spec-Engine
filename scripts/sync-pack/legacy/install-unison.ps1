# Install Unison inside WSL (interactive sudo).
# Run this once on each Windows machine before using sync-folder.ps1.

$ErrorActionPreference = 'Stop'

Write-Host 'Opening WSL to install Unison...' -ForegroundColor Cyan
Write-Host 'Enter your Linux password when sudo asks for it.' -ForegroundColor Yellow
Write-Host ''

wsl bash -lc "sudo apt-get update && sudo apt-get install -y unison openssh-client && echo '' && echo 'Done:' && unison -version | head -1"

if ($LASTEXITCODE -ne 0) {
    Write-Host 'Install failed. Open Ubuntu manually and run:' -ForegroundColor Red
    Write-Host '  sudo apt-get update && sudo apt-get install -y unison openssh-client'
    exit 1
}

Write-Host ''
Write-Host 'Unison is ready. Next: .\setup-sync.ps1' -ForegroundColor Green
