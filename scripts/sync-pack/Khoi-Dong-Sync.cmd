@echo off
chcp 65001 >nul
setlocal
title GoPaperless - Khoi dong Sync

set "ROOT=%~dp0"
set "WORKDIR="

if exist "%ROOT%scripts\sync-pack\khoi-dong-sync.ps1" (
    set "WORKDIR=%ROOT%scripts\sync-pack"
) else if exist "%ROOT%khoi-dong-sync.ps1" (
    set "WORKDIR=%ROOT%"
) else if exist "%ROOT%sync-pack\khoi-dong-sync.ps1" (
    set "WORKDIR=%ROOT%sync-pack"
)

if not defined WORKDIR (
    echo [LOI] Khong tim thay khoi-dong-sync.ps1
    pause
    exit /b 1
)

cd /d "%WORKDIR%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\khoi-dong-sync.ps1"
echo.
pause
exit /b %ERRORLEVEL%
