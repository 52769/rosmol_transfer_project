@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUNBUFFERED=1"
cd /d "%~dp0"
set "APP_DIR=%CD%\dist\ProjectTransfer"
if not exist ".venv\Scripts\python.exe" goto :install
goto :browser

:install
call "install.bat" --no-pause
if errorlevel 1 goto :fail

:browser
if not exist "%APP_DIR%" mkdir "%APP_DIR%"
set "PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\ms-playwright"
set "PLAYWRIGHT_SKIP_BROWSER_GC=1"
echo Installing Chromium to:
echo %PLAYWRIGHT_BROWSERS_PATH%
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto :fail
dir /s /b "%PLAYWRIGHT_BROWSERS_PATH%\chrome.exe" >nul 2>nul
if errorlevel 1 goto :missing
echo.
echo Chromium installed successfully.
echo Start the application through run_built_exe.bat.
pause
exit /b 0

:missing
echo.
echo ERROR: Chromium executable was not found after installation.
pause
exit /b 3

:fail
echo.
echo ERROR: Chromium installation failed.
pause
exit /b 1
