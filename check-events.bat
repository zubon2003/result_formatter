@echo off
rem ---------------------------------------------------------------------------
rem  FPVTrackside events/ JSON corruption checker (launcher).
rem  This file is ASCII only on purpose: cmd.exe re-reads the batch file with
rem  the active code page, so putting Japanese text here breaks parsing after
rem  "chcp 65001". All Japanese output lives in tools\check-events.ps1, which
rem  is UTF-8 with BOM and is read correctly by PowerShell either way.
rem ---------------------------------------------------------------------------
setlocal
chcp 65001 >nul

set "PS1=%~dp0tools\check-events.ps1"
if not exist "%PS1%" (
    echo [ERROR] not found: %PS1%
    echo Keep check-events.bat and tools\check-events.ps1 together.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "RC=%ERRORLEVEL%"

echo.
pause
exit /b %RC%
