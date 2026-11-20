@echo off
chcp 65001 >nul
title WeCom Group Export
set "ROOT=%~dp0"
set "LOG=%ROOT%runtime\export_error.log"
if not exist "%ROOT%runtime" mkdir "%ROOT%runtime" >nul 2>&1
if not exist "%ROOT%export_group.ps1" goto :missing
if not exist "%ROOT%.venv\Scripts\python.exe" goto :missing

powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%export_group.ps1" %* 2>>"%LOG%"
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" goto :failed
echo [OK] Finished. Press any key to close.
echo.
pause >nul
exit /b 0

:failed
echo [ERROR] The export tool exited with code %RC%.
echo Please check the log file: %LOG%
echo.
pause
exit /b %RC%

:missing
echo [ERROR] Required file not found under %ROOT%
echo Please make sure the repository files are intact.
echo.
pause
exit /b 1
