@echo off
setlocal
cd /d "%~dp0"
echo.
echo [sync-folder] Starting PowerShell...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync-folder.ps1" %*
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
    echo [sync-folder] FAILED with exit code %ERR%
) else (
    echo [sync-folder] Done.
)
exit /b %ERR%
