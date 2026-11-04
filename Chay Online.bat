@echo off
setlocal
set "APP_DIR=%~dp0"
if "%FLASK_RUN_PORT%"=="" set "FLASK_RUN_PORT=3000"

where cloudflared >nul 2>nul
if errorlevel 1 (
  echo Chua cai cloudflared. Cai bang lenh:  winget install Cloudflare.cloudflared
  echo Sau do bam lai file nay.
  pause
  exit /b 1
)

if exist "%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe" (
  set "APP_CMD="%APP_DIR%dist\QuanLyHoSoThanhNien\QuanLyHoSoThanhNien.exe""
) else if exist "%APP_DIR%.venv\Scripts\python.exe" (
  set "APP_CMD="%APP_DIR%.venv\Scripts\python.exe" "%APP_DIR%run.py""
) else (
  echo Khong tim thay ung dung. Chay build_windows.bat hoac tao .venv truoc.
  pause
  exit /b 1
)

rem Asks for a new admin password in this window when the admin still uses the default one.
%APP_CMD% --setup
if errorlevel 1 (
  pause
  exit /b 1
)

rem Only the tunnel on this computer may reach the app: it listens on 127.0.0.1, not on the LAN.
set "QLHS_ONLINE=1"
set "QLHS_HOST=127.0.0.1"
set "OPEN_BROWSER=0"
start "QLHS-app" /min cmd /c "%APP_CMD%"
timeout /t 5 /nobreak >nul

echo.
echo Dang mo duong ham. Tim dong "https://....trycloudflare.com" ben duoi va gui cho moi nguoi.
echo Khach tra cuu vao them /tra-cuu o cuoi dia chi.
echo.
cloudflared tunnel --url http://127.0.0.1:%FLASK_RUN_PORT%

taskkill /fi "WINDOWTITLE eq QLHS-app*" /t /f >nul 2>nul
