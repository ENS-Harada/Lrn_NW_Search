@echo off
chcp 65001 >nul 2>&1
title LRN Network Info Collector
echo.
echo  ================================================
echo    LRN Network Info Collector
echo    Starting...
echo  ================================================
echo.

set "PS_SCRIPT=%~dp0NW_Collect.ps1"

if not exist "%PS_SCRIPT%" (
    echo  [ERROR] NW_Collect.ps1 not found.
    echo  Place this bat file and NW_Collect.ps1 in the same folder.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%PS_SCRIPT%'"

echo.
pause
