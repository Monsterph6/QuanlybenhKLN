@echo off
cd /d "%~dp0"
echo Dang dong goi May chu (Service + Tray) thanh file .exe doc lap...
echo (khong can Python tren may dich sau khi dong goi xong)
echo.

if exist dist_server_raw rmdir /s /q dist_server_raw
if exist dist_server rmdir /s /q dist_server

python -m PyInstaller --noconfirm --onedir --console ^
    --hidden-import=win32timezone --contents-directory _internal_service ^
    --name QuanLyBenhNhanTHA-Service ^
    --distpath dist_server_raw --workpath build_server ^
    service.py
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --onedir --windowed ^
    --hidden-import=win32timezone --contents-directory _internal_tray ^
    --name QuanLyBenhNhanTHA-Tray ^
    --distpath dist_server_raw --workpath build_server ^
    server_tray.py
if errorlevel 1 goto :error

echo Dang gom Service + Tray vao chung 1 thu muc (dist_server\)...
mkdir dist_server
xcopy /e /i /y dist_server_raw\QuanLyBenhNhanTHA-Service\* dist_server\ >nul
xcopy /e /i /y dist_server_raw\QuanLyBenhNhanTHA-Tray\* dist_server\ >nul
copy /y install_server.bat dist_server\ >nul
copy /y uninstall_server.bat dist_server\ >nul
copy /y update_server.bat dist_server\ >nul
copy /y update_server.ps1 dist_server\ >nul
copy /y VERSION_SERVER.txt dist_server\ >nul
copy /y README.md dist_server\ >nul
rmdir /s /q dist_server_raw

echo.
echo Xong. Toan bo goi may chu nam o thu muc: dist_server\
echo Cai dat (van can quyen Administrator): dist_server\install_server.bat
pause
exit /b 0

:error
echo.
echo Dong goi that bai. Xem loi phia tren.
echo Kiem tra da chay "pip install -r requirements-server.txt pyinstaller" chua.
pause
exit /b 1
