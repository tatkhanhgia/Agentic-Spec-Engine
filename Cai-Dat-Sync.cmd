@echo off
chcp 65001 >nul
REM Chuyen sang project sync doc lap (ngang hang voi api-document-specification)
set "SYNC_PROJECT=%~dp0..\GoPaperless-Sync-Pack\Cai-Dat-Sync.cmd"
if exist "%SYNC_PROJECT%" (
    call "%SYNC_PROJECT%" %*
    exit /b %ERRORLEVEL%
)
REM Fallback: ban cu trong scripts\sync-pack (deprecated)
if exist "%~dp0scripts\sync-pack\Cai-Dat-Sync.cmd" (
    call "%~dp0scripts\sync-pack\Cai-Dat-Sync.cmd" %*
    exit /b %ERRORLEVEL%
)
echo.
echo [LOI] Khong tim thay GoPaperless-Sync-Pack.
echo Clone/copy project o: ..\GoPaperless-Sync-Pack
echo.
pause
exit /b 1
