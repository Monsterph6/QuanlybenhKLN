@echo off
cd /d "%~dp0"
echo ============================================================
echo  Go cai dat May chu chia se du lieu qua mang LAN
echo  (CAN CHAY VOI QUYEN ADMINISTRATOR)
echo ============================================================
echo.

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo LOI: Can chay file nay voi quyen Administrator.
    pause
    exit /b 1
)

set SERVICE_CMD=python service.py
if exist "%~dp0QuanLyBenhNhanTHA-Service.exe" set SERVICE_CMD="%~dp0QuanLyBenhNhanTHA-Service.exe"

%SERVICE_CMD% stop
%SERVICE_CMD% remove

echo.
echo Da dung va go cai dat dich vu.
echo Luu y: file benh_nhan.db va thu muc backups\ KHONG bi xoa - chi
echo xoa thu cong neu thuc su muon xoa het du lieu.
pause
