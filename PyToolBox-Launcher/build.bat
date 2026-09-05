@echo off
echo =======================================================
echo              PyToolBox-Launcher Build Script
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

echo [INFO] Compiling PyToolBox-Launcher.spec with PyInstaller...
python -m PyInstaller --noconfirm PyToolBox-Launcher.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] PyToolBox-Launcher Compilation failed!
    pause
    exit /b 1
)

echo [INFO] Compiling PyToolBox-Uninstaller.spec with PyInstaller...
python -m PyInstaller --noconfirm PyToolBox-Uninstaller.spec

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] PyToolBox-Uninstaller Compilation failed!
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
