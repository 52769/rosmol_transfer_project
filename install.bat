@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "NO_PAUSE=0"
if /I "%~1"=="--no-pause" set "NO_PAUSE=1"

set "PY_CMD="
set "PY_LABEL="

where py >nul 2>nul
if not errorlevel 1 (
    py -3.14 -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 14) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PY_CMD=py -3.14"
        set "PY_LABEL=Python 3.14"
    )
)

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3.13 -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 13) else 1)" >nul 2>nul
        if not errorlevel 1 (
            set "PY_CMD=py -3.13"
            set "PY_LABEL=Python 3.13"
        )
    )
)

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3.12 -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) else 1)" >nul 2>nul
        if not errorlevel 1 (
            set "PY_CMD=py -3.12"
            set "PY_LABEL=Python 3.12"
        )
    )
)

if not defined PY_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        python -c "import sys; raise SystemExit(0 if (3, 12) <= sys.version_info[:2] <= (3, 14) else 1)" >nul 2>nul
        if not errorlevel 1 (
            set "PY_CMD=python"
            set "PY_LABEL=Python from PATH"
        )
    )
)

if not defined PY_CMD goto :python_not_found

echo Selected %PY_LABEL%.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys; raise SystemExit(0 if (3, 12) <= sys.version_info[:2] <= (3, 14) else 1)" >nul 2>nul
    if errorlevel 1 (
        echo Removing broken or incompatible virtual environment...
        rmdir /s /q ".venv"
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment with %PY_LABEL%...
    %PY_CMD% -m venv ".venv"
    if errorlevel 1 goto :fail
)

echo Updating packaging tools...
".venv\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :fail

echo Installing application dependencies...
".venv\Scripts\python.exe" -m pip install -r "requirements_transfer.txt"
if errorlevel 1 goto :fail

echo Installing PyInstaller...
".venv\Scripts\python.exe" -m pip install --upgrade "pyinstaller>=6.21,<7" pyinstaller-hooks-contrib
if errorlevel 1 goto :fail

echo Installing Playwright Chromium...
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto :fail

echo.
".venv\Scripts\python.exe" --version
".venv\Scripts\python.exe" -m PyInstaller --version

echo Installation completed successfully.
if "%NO_PAUSE%"=="0" pause
exit /b 0

:python_not_found
echo.
echo ERROR: No working Python 3.12, 3.13, or 3.14 installation was found.
echo Run: py --list-paths
echo Run: py -3.14 --version
echo Disable broken Microsoft Store aliases or reinstall Python from python.org.
if "%NO_PAUSE%"=="0" pause
exit /b 2

:fail
echo.
echo ERROR: Installation failed.
echo Run: py --list-paths
echo Delete the .venv folder and run install.bat again.
if "%NO_PAUSE%"=="0" pause
exit /b 1
