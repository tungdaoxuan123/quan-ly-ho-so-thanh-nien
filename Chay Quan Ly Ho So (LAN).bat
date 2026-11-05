@echo off
setlocal
set "APP_DIR=%~dp0"

echo Mo ung dung cho cac may khac trong cung mang noi bo (LAN).
echo Moi nguoi phai dang nhap bang tai khoan do quan tri vien tao.
echo Quan tri vien phai doi mat khau mac dinh (admin123) truoc, bang cach chay
echo "Chay Quan Ly Ho So.bat" mot lan, dang nhap admin va doi mat khau.
set "QLHS_HOST=0.0.0.0"
echo.
echo Cua so ung dung se hien dia chi cho may khac (dang https://192.168.x.x:3000).
echo Neu Windows hoi ve tuong lua, tick ca Private va Public networks roi bam Allow access
echo (wifi khong co internet thuong la Public). Hoac chay "Mo tuong lua Windows (Run as administrator).bat".
echo.

rem Portable folder from build_windows_offline.sh: Python and libraries are bundled, nothing to install.
if exist "%APP_DIR%python\python.exe" (
  "%APP_DIR%python\python.exe" "%APP_DIR%run.py"
  echo.
  echo Ung dung da dung. Bam phim bat ky de dong cua so.
  pause >nul
  exit /b 0
)

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
