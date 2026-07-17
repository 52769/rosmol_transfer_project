@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" goto :install
".venv\Scripts\python.exe" -c "import PyInstaller" >nul 2>nul
if errorlevel 1 goto :install
goto :build

:install
call "install.bat" --no-pause
if errorlevel 1 goto :fail

:build
echo Cleaning old build output...
if exist "build" rmdir /s /q "build"
if exist "build_worker" rmdir /s /q "build_worker"
if exist "dist" rmdir /s /q "dist"
if exist "ProjectTransfer.spec" del /q "ProjectTransfer.spec"
if exist "ProjectTransferWorker.spec" del /q "ProjectTransferWorker.spec"

echo Building ProjectTransfer.exe...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed --name "ProjectTransfer" --collect-all playwright --hidden-import yaml --add-data "config_transfer.yaml;." "gui.py"
if errorlevel 1 goto :fail

echo Building ProjectTransferWorker.exe...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --console --name "ProjectTransferWorker" --distpath "dist\ProjectTransfer" --workpath "build_worker" --collect-all playwright --hidden-import yaml --hidden-import xlrd --hidden-import xlwt --hidden-import docx "worker.py"
if errorlevel 1 goto :fail

set "APP_DIR=%CD%\dist\ProjectTransfer"
set "APP_EXE=%APP_DIR%\ProjectTransfer.exe"
set "WORKER_EXE=%APP_DIR%\ProjectTransferWorker.exe"

if not exist "%APP_EXE%" goto :missing_exe
if not exist "%WORKER_EXE%" goto :missing_worker
if not exist "%APP_DIR%\_internal" goto :missing_runtime
dir /b "%APP_DIR%\_internal\python*.dll" >nul 2>nul
if errorlevel 1 goto :missing_runtime

".venv\Scripts\python.exe" "prepare_dist.py" "%APP_DIR%"
if errorlevel 1 goto :fail

echo Installing portable Playwright Chromium...
set "PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\ms-playwright"
set "PLAYWRIGHT_SKIP_BROWSER_GC=1"
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto :browser_fail
dir /s /b "%APP_DIR%\ms-playwright\chrome.exe" >nul 2>nul
if errorlevel 1 goto :browser_missing

rem Build contains temporary files and a non-distributable executable.
if exist "build" rmdir /s /q "build"
if exist "build_worker" rmdir /s /q "build_worker"

echo.
echo EXE created successfully.
echo Run only this file:
echo %APP_EXE%
echo.
echo Keep the complete dist\ProjectTransfer folder together.
start "" explorer.exe "%APP_DIR%"
pause
exit /b 0

:browser_fail
echo.
echo ERROR: Chromium download failed.
echo Check Internet access and run repair_playwright_browser.bat.
pause
exit /b 6

:browser_missing
echo.
echo ERROR: Playwright completed without a Chromium executable.
echo Run repair_playwright_browser.bat.
pause
exit /b 7

:missing_exe
echo.
echo ERROR: Build finished without ProjectTransfer.exe.
pause
exit /b 3

:missing_worker
echo.
echo ERROR: Build finished without ProjectTransferWorker.exe.
pause
exit /b 5

:missing_runtime
echo.
echo ERROR: Final runtime folder is incomplete.
echo Expected Python DLL under: %APP_DIR%\_internal
pause
exit /b 4

:fail
echo.
echo ERROR: EXE build failed.
pause
exit /b 1
