@echo off
cd /d "%~dp0"
echo Dang dong goi QuanLyBenhNhanTHA.exe ...
python -m PyInstaller --noconfirm --onedir --windowed ^
    --exclude-module numpy ^
    --name QuanLyBenhNhanTHA ^
    --distpath dist ^
    --workpath build ^
    app.py
if errorlevel 1 (
    echo.
    echo Dong goi that bai. Xem loi phia tren.
    pause
    exit /b 1
)
echo.
echo Xong. File chay o: dist\QuanLyBenhNhanTHA\QuanLyBenhNhanTHA.exe
pause
