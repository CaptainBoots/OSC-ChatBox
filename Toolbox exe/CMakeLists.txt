cmake_minimum_required(VERSION 4.2)
project(untitled)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_executable(untitled main.cpp untitled.rc)
set_target_properties(untitled PROPERTIES OUTPUT_NAME "ToolBox")

if(WIN32)
    set(APP_ICON_SVG "${CMAKE_SOURCE_DIR}/Boot's-ToolBox.svg")
    set(APP_ICON_ICO "${CMAKE_SOURCE_DIR}/Boot's-ToolBox-256.ico")

    add_custom_command(
        OUTPUT "${APP_ICON_ICO}"
        COMMAND powershell -NoProfile -ExecutionPolicy Bypass -File "${CMAKE_SOURCE_DIR}/make_icon.ps1" -SvgPath "${APP_ICON_SVG}" -IcoPath "${APP_ICON_ICO}"
        DEPENDS "${CMAKE_SOURCE_DIR}/make_icon.ps1" "${APP_ICON_SVG}"
        WORKING_DIRECTORY "${CMAKE_SOURCE_DIR}"
        COMMENT "Generate Boot's-ToolBox-256.ico from Boot's-ToolBox.svg"
        VERBATIM
    )

    add_custom_target(make_icon DEPENDS "${APP_ICON_ICO}")
    add_dependencies(untitled make_icon)
    set_source_files_properties(untitled.rc PROPERTIES OBJECT_DEPENDS "${APP_ICON_ICO}")
endif()

target_link_libraries(untitled PRIVATE winhttp ws2_32)

# No external system link libraries required.
# If your code needs specific Windows libraries, add them here.

# Prefer static CRT on MSVC so runtime DLLs are not required beside the exe
if(MSVC)
    set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
endif()

# For MinGW, prefer static linking to avoid runtime DLLs (libstdc++, libgcc, libwinpthread).
if(MINGW)
    # Statically link the stdc++/gcc runtimes and request static linking of system libs when possible.
    target_link_options(untitled PRIVATE -static -municode -static-libstdc++ -static-libgcc)

    # Detect MinGW bin/lib directories based on the compiler location.
    get_filename_component(MINGW_BIN_DIR "${CMAKE_CXX_COMPILER}" DIRECTORY)
    set(MINGW_LIB_DIR "${MINGW_BIN_DIR}/../lib")

    # Prefer a static winpthread if available.
    find_library(WINPTHREAD_STATIC NAMES winpthread libwinpthread PATHS "${MINGW_LIB_DIR}" NO_DEFAULT_PATH)
    if(WINPTHREAD_STATIC)
        message(STATUS "Found static winpthread: ${WINPTHREAD_STATIC}; linking statically.")
        target_link_options(untitled PRIVATE "-Wl,-Bstatic" "-lwinpthread" "-Wl,-Bdynamic")
    endif()
endif()
