@echo off
cd /d "%~dp0"
echo ============================================================
echo  Cai dat May chu chia se du lieu qua mang LAN
echo  (CAN CHAY VOI QUYEN ADMINISTRATOR - chuot phai file nay,
echo   chon "Run as administrator")
echo ============================================================
echo.

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo LOI: Can chay file nay voi quyen Administrator.
    echo Chuot phai vao install_server.bat, chon "Run as administrator".
    pause
    exit /b 1
)

if not exist lan_config.json (
    echo {"port": 8765}> lan_config.json
    echo Da tao lan_config.json voi cong mac dinh 8765.
    echo Muon doi cong: dung Notepad sua file lan_config.json roi chay lai file nay.
    echo.
)

echo Dang cai dat Windows Service...
python service.py --startup auto install
if errorlevel 1 goto :error

echo Dang bat dich vu...
python service.py start
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Da cai dat va bat dich vu thanh cong.
echo  Ten dich vu (xem qua services.msc): QuanLyBenhNhanTHA_Server
echo.
echo  De xem trang thai / dia chi IP tien loi hon (khong can quyen
echo  Administrator), chay server_tray.py - se hien icon o khay he
echo  thong. Co the them shortcut server_tray.py vao Startup de tu
echo  mo tray moi lan dang nhap (hoac tick "Khoi dong cung Windows"
echo  trong menu chuot phai cua icon tray).
echo ============================================================
pause
exit /b 0

:error
echo.
echo Cai dat that bai. Xem loi phia tren.
echo Kiem tra da chay "pip install -r requirements-server.txt" chua.
pause
exit /b 1
