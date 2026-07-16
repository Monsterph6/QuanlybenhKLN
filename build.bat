@echo off
cd /d "%~dp0"
echo Dang dong goi QuanLyBenhNhanTHA (ung dung chinh + thanh phan May chu) ...

if exist dist_raw rmdir /s /q dist_raw
if exist dist rmdir /s /q dist

python -m PyInstaller --noconfirm --onedir --windowed ^
    --exclude-module numpy ^
    --name QuanLyBenhNhanTHA ^
    --distpath dist_raw ^
    --workpath build ^
    app.py
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --onedir --console ^
    --hidden-import=win32timezone --exclude-module numpy --exclude-module openpyxl ^
    --contents-directory _internal_service ^
    --name QuanLyBenhNhanTHA-Service ^
    --distpath dist_raw --workpath build ^
    service.py
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --onedir --windowed ^
    --hidden-import=win32timezone --exclude-module numpy --exclude-module openpyxl ^
    --contents-directory _internal_tray ^
    --name QuanLyBenhNhanTHA-Tray ^
    --distpath dist_raw --workpath build ^
    server_tray.py
if errorlevel 1 goto :error

echo Dang gom tat ca vao 1 thu muc dist\QuanLyBenhNhanTHA\ ...
mkdir dist\QuanLyBenhNhanTHA
xcopy /e /i /y dist_raw\QuanLyBenhNhanTHA\* dist\QuanLyBenhNhanTHA\ >nul
xcopy /e /i /y dist_raw\QuanLyBenhNhanTHA-Service\* dist\QuanLyBenhNhanTHA\ >nul
xcopy /e /i /y dist_raw\QuanLyBenhNhanTHA-Tray\* dist\QuanLyBenhNhanTHA\ >nul
rmdir /s /q dist_raw

echo.
echo Xong. File chay o: dist\QuanLyBenhNhanTHA\QuanLyBenhNhanTHA.exe
echo (Cung thu muc co san QuanLyBenhNhanTHA-Service.exe / -Tray.exe cho vai tro May chu.)
pause
exit /b 0

:error
echo.
echo Dong goi that bai. Xem loi phia tren.
echo Kiem tra da chay "pip install -r requirements.txt pyinstaller" chua.
pause
exit /b 1
