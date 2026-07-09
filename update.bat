@echo off
cd /d "%~dp0"
echo Dang kiem tra ban cap nhat tu GitHub ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"
echo.
pause
