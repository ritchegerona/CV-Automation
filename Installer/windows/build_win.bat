@echo off
echo ==============================================
echo Building Windows Executable for CV Automation
echo ==============================================

:: Change directory to the project root (two levels up from this script's directory)
cd /d "%~dp0..\.."

:: Check if virtual environment exists
if not exist .venv (
    echo [ERROR] Python virtual environment (.venv) not found.
    echo Please create one first in the project root by running: python -m venv .venv
    exit /b 1
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Checking if PyInstaller is installed...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing now...
    pip install pyinstaller
)

echo Cleaning up previous builds...
if exist Installer\windows\build rmdir /s /q Installer\windows\build
if exist Installer\windows\dist rmdir /s /q Installer\windows\dist

echo Starting PyInstaller build process...
pyinstaller --workpath Installer\windows\build --distpath Installer\windows\dist --clean -y Installer\launcher.spec

echo ==============================================
echo Build finished! 
echo The standalone executable can be found at:
echo Installer\windows\dist\CV_Automation.exe
echo ==============================================
pause
