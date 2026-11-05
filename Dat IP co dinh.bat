@echo off
rem Pins the IP this PC has now, so the address phones use never changes. Asks for administrator rights itself.
net session >nul 2>&1
if not errorlevel 1 goto admin
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:admin
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ip_co_dinh.ps1"
echo.
echo Neu sau nay may doi sang wifi khac ma khong vao mang duoc, chay "Tra lai IP tu dong.bat".
pause
