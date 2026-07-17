@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUNBUFFERED=1"
cd /d "%~dp0"
set "APP_DIR=%CD%\dist\ProjectTransfer"
set "APP_EXE=%APP_DIR%\ProjectTransfer.exe"
set "WORKER_EXE=%APP_DIR%\ProjectTransferWorker.exe"
if not exist "%APP_EXE%" goto :missing
if not exist "%WORKER_EXE%" goto :incomplete
if not exist "%APP_DIR%\_internal" goto :incomplete
if exist "%APP_DIR%\ms-playwright" (
    set "PLAYWRIGHT_BROWSERS_PATH=%APP_DIR%\ms-playwright"
) else (
    set "PLAYWRIGHT_BROWSERS_PATH=%LOCALAPPDATA%\ms-playwright"
)
set "PLAYWRIGHT_SKIP_BROWSER_GC=1"
start "" "%APP_EXE%"
exit /b 0

:missing
echo ProjectTransfer.exe was not found.
echo Run build_exe.bat first.
pause
exit /b 2

:incomplete
echo The application folder is incomplete.
echo Keep the complete dist\ProjectTransfer folder together.
pause
exit /b 3
