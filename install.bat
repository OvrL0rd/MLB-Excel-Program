@echo off
setlocal enabledelayedexpansion

echo Checking for Python installation...

:: Try multiple methods to detect Python
python --version >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    goto :PYTHON_FOUND
)

py --version >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    goto :PYTHON_FOUND
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    goto :PYTHON_FOUND
)

echo Python is not installed. Downloading Python installer...

:: Download Python 3.11 installer (adjust verison as needed)
curl -o python-installer.exe https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe

if %ERRORLEVEL% NEQ 0 (
    echo Failed to download Python installer. Please check your internet connection or manually install Python.
    pause
    exit /b 1
)

echo Installing Python...
:: Install Python with pip, add to PATH, and disable path length limit
python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 Include_test=0

if $ERRORLEVEL% NEQ 0 (
    echo Failed to install Python.
    pause
    exit /b 1
)

del python_installer.exe

:: Refresh environment variables
call RefreshEnv.cmd

goto:INSTALL_PACKAGES

:PYTHON_FOUND
echo Python is already installed.

:INSTALL_PACKAGES
:: Install required Python packages
echo Installing required Python packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo Failed to install required Python packages. Please check the requirements.txt file.
    pause
    exit /b 1
)

echo Setup completed successfully.
echo You can now safely exit this window.
pause