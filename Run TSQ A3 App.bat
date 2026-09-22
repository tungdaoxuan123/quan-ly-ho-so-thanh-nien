@echo off
setlocal
set "APP_DIR=%~dp0"

if exist "%APP_DIR%dist\TSQ_A3\TSQ_A3.exe" (
  start "TSQ A3" "%APP_DIR%dist\TSQ_A3\TSQ_A3.exe"
  exit /b 0
)

if exist "%APP_DIR%.venv\Scripts\python.exe" (
  start "TSQ A3" "%APP_DIR%.venv\Scripts\python.exe" "%APP_DIR%tsq_a3_app.py"
  exit /b 0
)

echo TSQ_A3.exe was not found.
echo Build the Windows package with build_windows.bat, or install Python and create .venv.
pause
exit /b 1
