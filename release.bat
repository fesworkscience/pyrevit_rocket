@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File scripts\release.ps1 %*
pause
