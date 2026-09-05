@echo off
echo =======================================================
echo              SynthTool-Launcher C++ Build Script
echo =======================================================
echo.

:: 1. Search for Qt6 installation path
set "QT_PATH="
for /d %%d in (C:\Qt\6.*) do (
    if exist "%%d\msvc2022_64\bin\qmake.exe" (
        set "QT_PATH=%%d\msvc2022_64"
        goto qt_found
    )
)

:qt_found
if not "%QT_PATH%"=="" (
    echo [INFO] Auto-detected Qt6 at: %QT_PATH%
    goto configure_cmake
)

echo [WARNING] Qt6 MSVC 2022 64-bit installation was not auto-detected in C:\Qt\
echo.
echo Please enter the path to your Qt6 MSVC installation directory.
echo Example: C:\Qt\6.8.0\msvc2022_64
echo.
set /p "QT_PATH=Enter Qt6 path: "

if "%QT_PATH%"=="" (
    echo [ERROR] No Qt6 path specified. Cannot compile.
    pause
    exit /b 1
)

if not exist "%QT_PATH%\bin\qmake.exe" (
    echo [ERROR] Invalid Qt6 installation path! Could not find bin\qmake.exe.
    pause
    exit /b 1
)

:configure_cmake
echo.
echo [INFO] Configuring CMake with Qt6 path: %QT_PATH%
echo.

:: Check if cmake is available, if not, find JetBrains CMake
where cmake >nul 2>nul
if %errorlevel% neq 0 (
    if exist "C:\Program Files\JetBrains\CLion 2026.1.2\bin\cmake\win\x64\bin\cmake.exe" (
        set "PATH=%PATH%;C:\Program Files\JetBrains\CLion 2026.1.2\bin\cmake\win\x64\bin"
        echo [INFO] Added JetBrains CLion CMake to PATH.
    ) else (
        for /d %%d in ("C:\Program Files\JetBrains\CLion*") do (
            if exist "%%d\bin\cmake\win\x64\bin\cmake.exe" (
                set "PATH=%PATH%;%%d\bin\cmake\win\x64\bin"
                echo [INFO] Added %%d CLion CMake to PATH.
                goto cmake_set
            )
        )
    )
)
:cmake_set

:: Create build folder
if not exist "build" mkdir build
cd build

:: Run CMake config
cmake -G "Visual Studio 17 2022" -A x64 -DCMAKE_PREFIX_PATH="%QT_PATH%" ..
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] CMake configuration failed!
    cd ..
    pause
    exit /b 1
)

:: Run compilation
echo.
echo [INFO] Compiling SynthTool-Launcher...
echo.
cmake --build . --config Release
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Compilation failed!
    cd ..
    pause
    exit /b 1
)

echo.
echo =======================================================
echo [SUCCESS] Compilation completed successfully!
echo Executable is located in: build\Release\SynthTool-Launcher.exe
echo =======================================================
echo.
cd ..
pause
