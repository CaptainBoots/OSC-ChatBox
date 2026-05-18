
#include <windows.h>
#include <urlmon.h>
#include <filesystem>
#include <iostream>
#include <fstream>

#pragma comment(lib, "urlmon.lib")

namespace fs = std::filesystem;

const std::wstring PYTHON_VERSION = L"3.12.10";

const std::wstring PYTHON_INSTALLER_URL =
    L"https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe";

const std::wstring TOOLBOX_SCRIPT_URL =
    L"https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/main/VRChat-ToolBox.py";

bool FileExists(const std::wstring& path)
{
    return fs::exists(path);
}

std::wstring GetTempPathString()
{
    wchar_t buffer[MAX_PATH];
    GetTempPathW(MAX_PATH, buffer);
    return std::wstring(buffer);
}

bool DownloadFile(const std::wstring& url, const std::wstring& output)
{
    std::wcout << L"[Download] " << url << std::endl;

    HRESULT hr = URLDownloadToFileW(
        NULL,
        url.c_str(),
        output.c_str(),
        0,
        NULL
    );

    return SUCCEEDED(hr);
}

std::wstring FindPython()
{
    wchar_t buffer[4096];

    FILE* pipe = _wpopen(
        L"py -3 -c \"import sys; print(sys.executable)\"",
        L"r"
    );

    if (!pipe)
        return L"";

    std::wstring result;

    while (fgetws(buffer, 4096, pipe))
    {
        result += buffer;
    }

    _pclose(pipe);

    while (!result.empty() &&
          (result.back() == L'\n' || result.back() == L'\r'))
    {
        result.pop_back();
    }

    if (FileExists(result))
        return result;

    return L"";
}

bool InstallPython(const std::wstring& installerPath)
{
    std::wcout << L"[Python] Installing silently..." << std::endl;

    std::wstring cmd =
        L"\"" + installerPath + L"\" "
        L"/quiet InstallAllUsers=0 "
        L"PrependPath=1 "
        L"Include_pip=1 "
        L"SimpleInstall=1";

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {};

    wchar_t* cmdline = cmd.data();

    BOOL success = CreateProcessW(
        NULL,
        cmdline,
        NULL,
        NULL,
        FALSE,
        0,
        NULL,
        NULL,
        &si,
        &pi
    );

    if (!success)
        return false;

    WaitForSingleObject(pi.hProcess, INFINITE);

    DWORD exitCode = 1;
    GetExitCodeProcess(pi.hProcess, &exitCode);

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    return exitCode == 0;
}

bool RunPythonScript(
    const std::wstring& pythonExe,
    const std::wstring& scriptPath)
{
    std::wstring cmd =
        L"\"" + pythonExe + L"\" \"" + scriptPath + L"\"";

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {};

    wchar_t* cmdline = cmd.data();

    BOOL success = CreateProcessW(
        NULL,
        cmdline,
        NULL,
        NULL,
        FALSE,
        0,
        NULL,
        NULL,
        &si,
        &pi
    );

    if (!success)
        return false;

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    return true;
}

int wmain()
{
    std::wcout << L"==============================" << std::endl;
    std::wcout << L" VRChat ToolBox Bootstrapper " << std::endl;
    std::wcout << L"==============================" << std::endl;

    std::wstring pythonExe = FindPython();

    if (pythonExe.empty())
    {
        std::wcout << L"[Python] Not found." << std::endl;

        std::wstring tempDir = GetTempPathString();

        std::wstring installerPath =
            tempDir + L"python-installer.exe";

        if (!DownloadFile(
                PYTHON_INSTALLER_URL,
                installerPath))
        {
            MessageBoxW(
                NULL,
                L"Failed to download Python installer.",
                L"Error",
                MB_ICONERROR
            );

            return 1;
        }

        if (!InstallPython(installerPath))
        {
            MessageBoxW(
                NULL,
                L"Python installation failed.",
                L"Error",
                MB_ICONERROR
            );

            return 1;
        }

        pythonExe = FindPython();

        if (pythonExe.empty())
        {
            MessageBoxW(
                NULL,
                L"Python installed but could not be found.",
                L"Error",
                MB_ICONERROR
            );

            return 1;
        }
    }

    std::wcout << L"[Python] Found: " << pythonExe << std::endl;

    std::wstring toolboxPath =
        GetTempPathString() + L"VRChat-ToolBox.py";

    if (!DownloadFile(
            TOOLBOX_SCRIPT_URL,
            toolboxPath))
    {
        MessageBoxW(
            NULL,
            L"Failed to download toolbox script.",
            L"Error",
            MB_ICONERROR
        );

        return 1;
    }

    std::wcout << L"[ToolBox] Downloaded." << std::endl;

    if (!RunPythonScript(pythonExe, toolboxPath))
    {
        MessageBoxW(
            NULL,
            L"Failed to launch toolbox script.",
            L"Error",
            MB_ICONERROR
        );

        return 1;
    }

    std::wcout << L"[ToolBox] Started successfully." << std::endl;

    return 0;
}

