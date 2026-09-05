cmake_minimum_required(VERSION 3.16)

# 1. Dynamically read version from the Python ProtoTool-Launcher.py as single-source-of-truth
set(LAUNCHER_VERSION "1.0.1") # Fallback default
set(PY_FILE "${CMAKE_CURRENT_SOURCE_DIR}/../ProtoTool-Launcher/ProtoTool-Launcher.py")

if(EXISTS "${PY_FILE}")
    file(READ "${PY_FILE}" PY_CONTENT)
    if(PY_CONTENT MATCHES "VERSION\\s*=\\s*\"([^\"]+)\"")
        set(LAUNCHER_VERSION "${CMAKE_MATCH_1}")
        message(STATUS "Auto-detected ProtoTool-Launcher version: ${LAUNCHER_VERSION}")
    elseif(PY_CONTENT MATCHES "VERSION\\s*=\\s*'([^']+)'")
        set(LAUNCHER_VERSION "${CMAKE_MATCH_1}")
        message(STATUS "Auto-detected ProtoTool-Launcher version: ${LAUNCHER_VERSION}")
    endif()
endif()

project(SynthTool-Launcher VERSION ${LAUNCHER_VERSION} LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

# Find Qt6 modules
find_package(Qt6 REQUIRED COMPONENTS Core Gui Widgets Network)

# Pass the dynamic version to the C++ compiler as a preprocessor definition
add_compile_definitions(APP_VERSION="${PROJECT_VERSION}")

add_executable(SynthTool-Launcher
    src/main.cpp
    src/mainwindow.h
    src/mainwindow.cpp
    src/onboardingwizard.h
    src/onboardingwizard.cpp
    src/settingsdialog.h
    src/settingsdialog.cpp
    src/helpdialog.h
    src/helpdialog.cpp
    src/configmanager.h
    src/configmanager.cpp
    src/customwidgets.h
    src/customwidgets.cpp
    src/theme.h
    src/theme.cpp
)

target_link_libraries(SynthTool-Launcher PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
    Qt6::Network
)

# Add Uninstaller executable
add_executable(SynthTool-Uninstaller
    src/uninstaller.cpp
)

target_link_libraries(SynthTool-Uninstaller PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
)

# On Windows, build as a windowed application (hides the console window)
if(WIN32)
    set_target_properties(SynthTool-Launcher PROPERTIES WIN32_EXECUTABLE TRUE)
    set_target_properties(SynthTool-Uninstaller PROPERTIES WIN32_EXECUTABLE TRUE)
endif()
