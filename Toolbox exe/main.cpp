#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <winhttp.h>
#include <filesystem>
#include <iostream>
#include <cstdio>
#include <string_view>
#include <vector>
#include <algorithm>

namespace fs = std::filesystem;

inline constexpr std::wstring_view TOOLBOX_SCRIPT_URL =
    L"https://raw.githubusercontent.com/CaptainBoots/VRChat-ToolBox/main/VRChat-ToolBox.py";
inline constexpr std::wstring_view PYTHON_VERSION = L"3.12.10";
inline constexpr std::wstring_view APP_STATE_DIR_NAME = L"VRChat-ToolBox";
inline constexpr std::wstring_view FIRST_RUN_MARKER_NAME = L"first-run-complete.flag";

std::wstring GetEnvironmentVariableString(const wchar_t* name)
{
    DWORD required = GetEnvironmentVariableW(name, nullptr, 0);
    if (required == 0) return L"";

    std::wstring value(required, L'\0');
    DWORD written = GetEnvironmentVariableW(name, value.data(), required);
    if (written == 0 || written >= required) return L"";

    value.resize(written);
    return value;
}

std::wstring GetAppStateDirString()
{
    std::wstring localAppData = GetEnvironmentVariableString(L"LOCALAPPDATA");
    if (localAppData.empty()) return L"";

    fs::path stateDir = fs::path(localAppData) / std::wstring(APP_STATE_DIR_NAME);
    std::error_code ec;
    fs::create_directories(stateDir, ec);
    if (ec) return L"";

    return stateDir.wstring();
}

std::wstring GetFirstRunMarkerPath()
{
    std::wstring stateDir = GetAppStateDirString();
    if (stateDir.empty()) return L"";

    return (fs::path(stateDir) / std::wstring(FIRST_RUN_MARKER_NAME)).wstring();
}

bool IsFirstRun()
{
    std::wstring marker = GetFirstRunMarkerPath();
    if (marker.empty()) return true;
    return !fs::exists(marker);
}

void MarkFirstRunComplete()
{
    std::wstring marker = GetFirstRunMarkerPath();
    if (marker.empty()) return;

    HANDLE hFile = CreateFileW(
        marker.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_HIDDEN,
        nullptr
    );

    if (hFile != INVALID_HANDLE_VALUE)
        CloseHandle(hFile);
}

std::wstring GetTempPathString()
{
    wchar_t buffer[MAX_PATH];
    DWORD len = GetTempPathW(MAX_PATH, buffer);
    if (len == 0 || len > MAX_PATH) return L"";
    return std::wstring(buffer);
}

bool FileExists(const std::wstring& path)
{
    return fs::exists(path);
}

std::wstring GetExeDirString()
{
    wchar_t buffer[MAX_PATH];
    GetModuleFileNameW(nullptr, buffer, MAX_PATH);
    fs::path p(buffer);
    return p.parent_path().wstring();
}

bool DownloadFile(std::wstring_view url, const std::wstring& output)
{
    std::wstring urlStr(url);
    size_t pos = urlStr.find(L"://");
    if (pos == std::wstring::npos) return false;
    std::wstring scheme = urlStr.substr(0, pos);
    bool secure = (_wcsicmp(scheme.c_str(), L"https") == 0);
    size_t hostStart = pos + 3;
    size_t pathStart = urlStr.find(L'/', hostStart);
    std::wstring host = (pathStart == std::wstring::npos) ? urlStr.substr(hostStart) : urlStr.substr(hostStart, pathStart - hostStart);
    std::wstring path = (pathStart == std::wstring::npos) ? L"/" : urlStr.substr(pathStart);

    HINTERNET hSession = WinHttpOpen(L"VRChat-ToolBox-Downloader/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME,
                                     WINHTTP_NO_PROXY_BYPASS,
                                     0);
    if (!hSession) return false;

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), secure ? INTERNET_DEFAULT_HTTPS_PORT : INTERNET_DEFAULT_HTTP_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

    DWORD flags = secure ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    BOOL sent = WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0);
    if (!sent) { WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    BOOL received = WinHttpReceiveResponse(hRequest, NULL);
    if (!received) { WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    HANDLE hFile = CreateFileW(output.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) { WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    const DWORD bufSize = 8192;
    std::vector<char> buffer(bufSize);
    DWORD dwSize = 0;

    do {
        if (!WinHttpQueryDataAvailable(hRequest, &dwSize)) break;
        if (dwSize == 0) break;
        DWORD toRead = dwSize;
        while (toRead > 0) {
            DWORD chunk = std::min<DWORD>(toRead, bufSize);
            DWORD bytesRead = 0;
            if (!WinHttpReadData(hRequest, buffer.data(), chunk, &bytesRead) || bytesRead == 0) {
                CloseHandle(hFile);
                WinHttpCloseHandle(hRequest);
                WinHttpCloseHandle(hConnect);
                WinHttpCloseHandle(hSession);
                return false;
            }
            DWORD written = 0;
            if (!WriteFile(hFile, buffer.data(), bytesRead, &written, NULL) || written != bytesRead) {
                CloseHandle(hFile);
                WinHttpCloseHandle(hRequest);
                WinHttpCloseHandle(hConnect);
                WinHttpCloseHandle(hSession);
                return false;
            }
            toRead -= bytesRead;
        }
    } while (dwSize > 0);

    CloseHandle(hFile);
    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    return true;
}

std::wstring BuildPythonInstallerUrl()
{
    std::wstring version(PYTHON_VERSION);
    return L"https://www.python.org/ftp/python/" + version + L"/python-" + version + L"-amd64.exe";
}

std::wstring GetPythonMajorMinorTag()
{
    std::wstring version(PYTHON_VERSION);
    const size_t firstDot = version.find(L'.');
    if (firstDot == std::wstring::npos) return L"";

    const size_t secondDot = version.find(L'.', firstDot + 1);
    if (secondDot == std::wstring::npos) return L"";

    return version.substr(0, firstDot) + version.substr(firstDot + 1, secondDot - firstDot - 1);
}

std::wstring GetDefaultInstalledPythonPath()
{
    std::wstring localAppData = GetEnvironmentVariableString(L"LOCALAPPDATA");
    std::wstring majorMinorTag = GetPythonMajorMinorTag();
    if (localAppData.empty() || majorMinorTag.empty()) return L"";

    std::wstring versionFolder = L"Python" + majorMinorTag;
    std::vector<fs::path> candidates = {
        fs::path(localAppData) / L"Programs" / L"Python" / versionFolder / L"python.exe",
        fs::path(localAppData) / L"Programs" / L"Python" / (versionFolder + L"-64") / L"python.exe"
    };

    for (const fs::path& candidate : candidates)
    {
        if (fs::exists(candidate))
            return candidate.wstring();
    }

    return L"";
}

bool RunProcessAndWait(const std::wstring& commandLine)
{
    STARTUPINFOW si{ sizeof(si) };
    PROCESS_INFORMATION pi{};

    std::wstring cmdCopy = commandLine;
    wchar_t* cmdline = cmdCopy.data();

    BOOL success = CreateProcessW(
        nullptr,
        cmdline,
        nullptr,
        nullptr,
        FALSE,
        0,
        nullptr,
        nullptr,
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

bool InstallPythonForCurrentUser()
{
    const std::wstring tempDir = GetTempPathString();
    if (tempDir.empty())
        return false;

    const std::wstring installerName = L"python-" + std::wstring(PYTHON_VERSION) + L"-amd64.exe";
    const std::wstring installerPath = (fs::path(tempDir) / installerName).wstring();
    const std::wstring installerUrl = BuildPythonInstallerUrl();

    if (!DownloadFile(installerUrl, installerPath))
        return false;

    std::wstring command =
        L"\"" + installerPath + L"\" "
        L"/quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1";

    const bool ok = RunProcessAndWait(command);
    DeleteFileW(installerPath.c_str());
    return ok;
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

bool RunPythonScript(const std::wstring& pythonExe, const std::wstring& scriptPath)
{
    std::wstring cmd =
        L"\"" + pythonExe + L"\" \"" + scriptPath + L"\"";

    STARTUPINFOW si{ sizeof(si) };
    PROCESS_INFORMATION pi{};

    std::wstring cmdCopy = cmd;
    wchar_t* cmdline = cmdCopy.data();

    BOOL success = CreateProcessW(
        nullptr,
        cmdline,
        nullptr,
        nullptr,
        FALSE,
        0,
        nullptr,
        nullptr,
        &si,
        &pi
    );

    if (!success)
        return false;

    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    return true;
}

int wmain(int argc, wchar_t* argv[])
{
    std::wcout << L"==============================" << std::endl;
    std::wcout << L"        VRChat ToolBox        " << std::endl;
    std::wcout << L"==============================" << std::endl;

    bool firstRun = IsFirstRun();
    std::wstring pythonExe = FindPython();

    if (pythonExe.empty() && firstRun)
    {
        std::wcout << L"[Python] Not found. First run detected, installing Python " << PYTHON_VERSION << L"..." << std::endl;

        if (!InstallPythonForCurrentUser())
        {
            MessageBoxW(
                nullptr,
                L"Failed to install Python automatically on first run.",
                L"Python Install Failed",
                MB_ICONERROR
            );
            MarkFirstRunComplete();
            return 1;
        }

        pythonExe = FindPython();
        if (pythonExe.empty())
            pythonExe = GetDefaultInstalledPythonPath();

        MarkFirstRunComplete();
    }
    else if (firstRun)
    {
        MarkFirstRunComplete();
    }

    if (pythonExe.empty())
    {
        MessageBoxW(
            nullptr,
            L"Python (3.x) not found. Install Python and ensure the 'py' launcher or python executable is available.",
            L"Missing Python",
            MB_ICONERROR
        );
        return 1;
    }

    std::wcout << L"[Python] Found: " << pythonExe << std::endl;

    std::wstring toolboxPath;

    if (argc > 1)
    {
        toolboxPath = argv[1];
    }
    else
    {
        fs::path p(GetExeDirString());
        toolboxPath = (p / L"VRChat-ToolBox.py").wstring();
    }

    if (!FileExists(toolboxPath))
    {
        // Try to download the toolbox script to the temp directory.
        std::wstring tempDir = GetTempPathString();
        if (tempDir.empty()) {
            MessageBoxW(nullptr, L"Failed to determine temp directory.", L"Error", MB_ICONERROR);
            return 1;
        }
        std::wstring downloadPath = (fs::path(tempDir) / L"VRChat-ToolBox.py").wstring();
        if (!DownloadFile(TOOLBOX_SCRIPT_URL, downloadPath)) {
            MessageBoxW(nullptr, L"Failed to download toolbox script.", L"Error", MB_ICONERROR);
            return 1;
        }
        toolboxPath = downloadPath;
    }

    if (!RunPythonScript(pythonExe, toolboxPath))
    {
        MessageBoxW(
            nullptr,
            L"Failed to launch toolbox script.",
            L"Error",
            MB_ICONERROR
        );
        return 1;
    }

    return 0;
}
