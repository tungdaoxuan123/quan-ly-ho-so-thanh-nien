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

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed --name TSQ_A3 --add-data "TSQ_A3_Full_Template.docx;." tsq_a3_app.py
if errorlevel 1 exit /b 1

echo Built dist\TSQ_A3\TSQ_A3.exe
endlocal
