@echo off
echo =======================================================
echo              ProtoTool-Launcher Build Script
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

echo [INFO] Compiling ProtoTool-Launcher.spec with PyInstaller...
python -m PyInstaller --noconfirm ProtoTool-Launcher.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] ProtoTool-Launcher Compilation failed!
    pause
    exit /b 1
)

echo [INFO] Compiling Uninstaller.py with PyInstaller...
python -m PyInstaller --onefile --name="Uninstaller" Uninstaller.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Uninstaller Compilation failed!
    pause
    exit /b 1
)

echo.
echo =======================================================
echo [SUCCESS] Compilation completed!
echo Executables are located in: dist\
echo =======================================================
echo.
pause
