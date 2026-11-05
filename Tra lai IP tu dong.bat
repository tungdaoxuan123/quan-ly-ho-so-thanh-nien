@echo off
rem Undoes "Dat IP co dinh.bat": the router hands out the address again. Asks for administrator rights itself.
net session >nul 2>&1
if not errorlevel 1 goto admin
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:admin
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ip_co_dinh.ps1" -Undo
pause
