@echo off
rem Sidenote setup. Double-click to run.
chcp 65001 > nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
echo.
pause
