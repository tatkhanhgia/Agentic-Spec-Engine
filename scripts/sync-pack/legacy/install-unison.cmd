@echo off
setlocal
cd /d "%~dp0"
echo.
echo [install-unison] Starting PowerShell...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-unison.ps1"
set ERR=%ERRORLEVEL%
echo.
if %ERR% neq 0 (
    echo [install-unison] FAILED with exit code %ERR%
) else (
    echo [install-unison] Done.
)
exit /b %ERR%
