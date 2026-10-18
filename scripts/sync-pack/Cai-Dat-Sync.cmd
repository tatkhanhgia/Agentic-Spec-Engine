@echo off
chcp 65001 >nul
setlocal
title GoPaperless - Cai dat Sync

set "ROOT=%~dp0"
set "WORKDIR="

if exist "%ROOT%scripts\sync-pack\install.ps1" (
    set "WORKDIR=%ROOT%scripts\sync-pack"
) else if exist "%ROOT%install.ps1" (
    set "WORKDIR=%ROOT%"
) else if exist "%ROOT%sync-pack\install.ps1" (
    set "WORKDIR=%ROOT%sync-pack"
)

if not defined WORKDIR (
    echo.
    echo [LOI] Khong tim thay install.ps1
    echo.
    echo Dam bao ban chay file nay trong thu muc da giai nen,
    echo hoac trong thu muc api-document-specification.
    echo.
    echo Thu muc hien tai: %ROOT%
    echo.
    pause
    exit /b 1
)

cd /d "%WORKDIR%"

echo.
echo ========================================
echo    GO PAPERLESS - BO CAI SYNC FILE
echo ========================================
echo.
echo Thu muc: %WORKDIR%
echo.
echo Doc README.txt neu ban moi bat dau.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\install.ps1" -Mode Syncthing %*
set ERR=%ERRORLEVEL%

echo.
if %ERR% neq 0 (
    echo [LOI] Cai dat that bai. Ma loi: %ERR%
) else (
    echo [OK] Hoan tat.
)
echo.
pause
exit /b %ERR%
