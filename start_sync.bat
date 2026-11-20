@echo off
chcp 65001 >nul
title WeCom Sync - Start
set "ROOT=%~dp0"
set "PIDFILE=%ROOT%runtime\sync.pid"
if not exist "%ROOT%runtime" mkdir "%ROOT%runtime" >nul

if exist "%PIDFILE%" (
  for /f "usebackq tokens=* delims=" %%P in ("%PIDFILE%") do (
    tasklist /FI "PID eq %%P" 2>nul | find /I "powershell" >nul
    if not errorlevel 1 (
      echo [INFO] Background sync is already running, PID=%%P
      echo Use stop_sync.bat first if you want to restart it.
      goto :end
    )
  )
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $root = '%ROOT%'; $p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File',($root + 'sync_watch.ps1')) -WindowStyle Hidden -PassThru; Set-Content -Path ($root + 'runtime\sync.pid') -Value $p.Id -Encoding ASCII; Write-Host ('[OK] Background sync started (hidden window), PID=' + $p.Id) }"
if errorlevel 1 (
  echo [ERROR] Failed to start. See runtime\sync.log
) else (
  echo [OK] Started. Log: %ROOT%runtime\sync.log
)

:end
echo.
pause
