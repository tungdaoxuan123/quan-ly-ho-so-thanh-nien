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

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed --name QuanLyHoSoThanhNien --paths src --add-data "Mau_Ho_So_Thanh_Nien.docx;." --add-data "Mau_Danh_Sach_Thanh_Nien.xlsx;." run.py
if errorlevel 1 exit /b 1

echo Built dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe
endlocal
