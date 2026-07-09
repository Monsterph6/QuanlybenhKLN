@echo off
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo Da xay ra loi khi chay ung dung. Xem thong bao loi phia tren.
    pause
)
