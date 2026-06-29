@echo off
REM One-command build of the Axon installer. Double-click this, or run: build.bat
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
pause
