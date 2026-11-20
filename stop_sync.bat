@echo off
chcp 65001 >nul
title WeCom Sync - Stop
set "PIDFILE=%~dp0runtime\sync.pid"

if not exist "%PIDFILE%" (
  echo [INFO] No running sync process found (sync.pid missing).
  goto :end
)

for /f "usebackq tokens=* delims=" %%P in ("%PIDFILE%") do (
  tasklist /FI "PID eq %%P" 2>nul | find /I "powershell" >nul
  if not errorlevel 1 (
    taskkill /PID %%P /T /F >nul 2>&1
    echo [OK] Stopped background sync, PID=%%P
  ) else (
    echo [INFO] PID %%P is no longer running; cleaned up the record.
  )
)
del "%PIDFILE%" >nul 2>&1

:end
echo.
pause
