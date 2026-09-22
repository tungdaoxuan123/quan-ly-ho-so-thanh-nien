@echo off
setlocal
set "APP_DIR=%~dp0"

if exist "%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe" (
  start "Quan ly ho so thanh nien" "%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe"
  exit /b 0
)

if exist "%APP_DIR%.venv\Scripts\python.exe" (
  start "Quan ly ho so thanh nien" "%APP_DIR%.venv\Scripts\python.exe" "%APP_DIR%quan_ly_ho_so.py"
  exit /b 0
)

echo QuanLyHoSoThanhNien.exe was not found.
echo Build the Windows package with build_windows.bat, or install Python and create .venv.
pause
exit /b 1
