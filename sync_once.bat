@echo off
chcp 65001 >nul
title WeCom Sync - Run Once
set "ROOT=%~dp0"
set "PY=%ROOT%.venv\Scripts\python.exe"
if not exist "%PY%" goto :missing

echo Forcing one sync for the groups in watch_groups.txt ...
echo.
"%PY%" -X utf8 "%ROOT%wecom_tools.py" sync --force
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
  echo [ERROR] Sync failed with code %RC%.
) else (
  echo [OK] Done. Background details: %ROOT%runtime\sync.log
)
echo.
pause
exit /b %RC%

:missing
echo [ERROR] Python not found: %PY%
echo.
pause
exit /b 1
