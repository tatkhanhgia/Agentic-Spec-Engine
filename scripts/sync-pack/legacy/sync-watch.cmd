@echo off
setlocal
cd /d "%~dp0"
echo.
echo [sync-watch] Continuous bidirectional sync. Press Ctrl+C to stop.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync-folder.ps1" -Watch
exit /b %ERRORLEVEL%
