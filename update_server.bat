@echo off
cd /d "%~dp0"
echo Dang kiem tra ban cap nhat MAY CHU tu GitHub ...
echo (Can quyen Administrator - neu bao loi quyen, chuot phai file nay
echo  va chon "Run as administrator")
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_server.ps1"
echo.
pause
