@echo off
cd /d "%~dp0"
echo Dang dong goi May chu (Service + Tray) thanh file .exe doc lap...
echo (khong can Python tren may dich sau khi dong goi xong)
echo.

python -m PyInstaller --noconfirm --onedir --console ^
    --hidden-import=win32timezone ^
    --name QuanLyBenhNhanTHA-Service ^
    --distpath dist_server --workpath build_server ^
    service.py
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --onedir --windowed ^
    --hidden-import=win32timezone ^
    --name QuanLyBenhNhanTHA-Tray ^
    --distpath dist_server --workpath build_server ^
    server_tray.py
if errorlevel 1 goto :error

echo.
echo Xong. File chay o:
echo   dist_server\QuanLyBenhNhanTHA-Service\QuanLyBenhNhanTHA-Service.exe
echo   dist_server\QuanLyBenhNhanTHA-Tray\QuanLyBenhNhanTHA-Tray.exe
echo.
echo Cai dat dich vu tu ban da dong goi (van can quyen Administrator):
echo   QuanLyBenhNhanTHA-Service.exe --startup auto install
echo   QuanLyBenhNhanTHA-Service.exe start
pause
exit /b 0

:error
echo.
echo Dong goi that bai. Xem loi phia tren.
echo Kiem tra da chay "pip install -r requirements-server.txt pyinstaller" chua.
pause
exit /b 1
