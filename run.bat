@echo off
setlocal

cd /d "%~dp0"

echo ==========================================
echo SAP Favorites Automation
echo ==========================================
echo.
set /p WINDOWKEY=Enter SAP window title keyword (example: WAP or WAP(1)/400): 

if "%WINDOWKEY%"=="" (
    echo.
    echo No keyword entered. Using default: WAP
    set WINDOWKEY=WAP
)

echo.
echo Running script with window title keyword: %WINDOWKEY%
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0main.ps1" -WindowTitleKeyword "%WINDOWKEY%"

echo.
pause