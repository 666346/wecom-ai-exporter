@echo off
chcp 65001 >nul
title WeCom AI Export Tool - Local Web UI
set "ROOT=%~dp0"
set "PY=%ROOT%.venv\Scripts\python.exe"
set "LOG=%ROOT%runtime\web_error.log"
if not exist "%ROOT%runtime" mkdir "%ROOT%runtime" >nul 2>&1
if not exist "%PY%" goto :missing
if not exist "%ROOT%web_app.py" goto :missing
if not exist "%ROOT%web_ui.html" goto :missing

echo.
echo   WeCom AI Export Tool
echo   ---------------------------------------------
echo   Local service : http://127.0.0.1:8765
echo   The browser will open automatically.
echo   KEEP THIS WINDOW OPEN. Close it to stop.
echo.
"%PY%" -X utf8 "%ROOT%web_app.py" 2>>"%LOG%"
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
  echo [ERROR] Service exited with code %RC%.
  echo Check the log file: %LOG%
) else (
  echo [INFO] Service stopped normally.
)
echo.
pause
exit /b %RC%

:missing
echo [ERROR] Required file not found under %ROOT%
echo Please make sure the repository files are intact.
echo.
pause
exit /b 1
