@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUNBUFFERED=1"
cd /d "%~dp0"
set "MODE="
set "ROW="
set /p MODE=Mode, download or upload: 
set /p ROW=XLS row number: 
if not exist ".venv\Scripts\python.exe" call "install.bat" --no-pause
if errorlevel 1 goto :fail
".venv\Scripts\python.exe" "optimized_transfer.py" --config "config_transfer.yaml" --mode "%MODE%" --row "%ROW%"
set "CODE=%ERRORLEVEL%"
pause
exit /b %CODE%
:fail
echo ERROR: Application environment is not ready.
pause
exit /b 1
