@echo off
chcp 65001 >nul
title PDF Giao An - May chu chuyen doi PDF Toan sang Word

set "PATH=C:\Users\TranN\AppData\Local\poppler\poppler-26.09.0\Library\bin;C:\Users\TranN\AppData\Local\Pandoc;C:\Users\TranN\AppData\Local\Programs\Python\Python313;C:\Users\TranN\AppData\Local\Programs\Python\Python313\Scripts;%PATH%"
set "PYTHONUTF8=1"

cd /d "%~dp0"

echo =======================================================
echo          DANG KHOI DONG PDF GIAO AN...
echo =======================================================
echo Ung dung se tu dong mo tren trinh duyet tai:
echo http://localhost:8000
echo.
echo Nhan Ctrl+C de dung may chu khi khong su dung.
echo =======================================================

start "" http://localhost:8000
py -3.13 app.py
if errorlevel 1 (
    python app.py
)
pause
