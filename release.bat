@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File release_scripts\release.ps1 %*
pause
