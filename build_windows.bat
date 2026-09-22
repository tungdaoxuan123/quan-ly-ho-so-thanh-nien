@echo off
setlocal
set "APP_DIR=%~dp0"
cd /d "%APP_DIR%"

if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
  if errorlevel 1 exit /b 1
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 exit /b 1

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed --name QuanLyHoSoThanhNien --add-data "Mau_Ho_So_Thanh_Nien.docx;." quan_ly_ho_so.py
if errorlevel 1 exit /b 1

echo Built dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe
endlocal
