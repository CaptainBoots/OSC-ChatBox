@echo off
echo =======================================================
echo              VRChat ToolBox Build Script
echo =======================================================
echo.

:: Ensure pip is installed
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python/pip is not installed or not in PATH!
    echo Please install Python and try again.
    pause
    exit /b 1
)

:: Ensure PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% equ 0 goto pyinstaller_installed

echo [INFO] PyInstaller is not installed. Installing now...
python -m pip install pyinstaller
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install PyInstaller!
    pause
    exit /b 1
)

:pyinstaller_installed

echo [INFO] Compiling VRChat-ToolBox.py with PyInstaller...
python -m PyInstaller --noconsole --onefile --icon="Images/Boot's-ToolBox.ico" --name="ToolBox" VRChat-ToolBox.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Compilation failed!
    pause
    exit /b 1
)

echo.
echo =======================================================
echo [SUCCESS] Compilation completed!
echo Executable is located at: dist\ToolBox.exe
echo =======================================================
echo.
pause
