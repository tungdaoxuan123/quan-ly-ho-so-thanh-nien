@echo off
setlocal
set "APP_DIR=%~dp0"

echo Mo ung dung cho cac may khac trong cung mang noi bo (LAN).
echo Moi may phai nhap mat khau nay moi vao duoc.
set /p "QLHS_PASSWORD=Dat mat khau: "
if "%QLHS_PASSWORD%"=="" (
  echo Can nhap mat khau.
  pause
  exit /b 1
)
set "QLHS_HOST=0.0.0.0"
echo.
echo Cua so ung dung se hien dia chi cho may khac (dang http://192.168.x.x:8765).
echo Neu Windows hoi ve tuong lua, chon Private networks va bam Allow access.
echo.

if exist "%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe" (
  "%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe"
  exit /b 0
)

if exist "%APP_DIR%.venv\Scripts\python.exe" (
  "%APP_DIR%.venv\Scripts\python.exe" "%APP_DIR%run.py"
  exit /b 0
)

echo QuanLyHoSoThanhNien.exe was not found.
echo Build the Windows package with build_windows.bat, or install Python and create .venv.
pause
exit /b 1
