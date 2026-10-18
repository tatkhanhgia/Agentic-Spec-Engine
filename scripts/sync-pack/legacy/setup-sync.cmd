@echo off
setlocal
cd /d "%~dp0"
echo.
echo [setup-sync] Starting PowerShell...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup-sync.ps1" %*
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
    echo [setup-sync] FAILED with exit code %ERR%
) else (
    echo [setup-sync] Done.
)
exit /b %ERR%
