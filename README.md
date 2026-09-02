# ProtoTool-Launcher

<img src="Images/Boot%27s-ToolBox.svg" alt="Boot's ToolBox" width="200" />

Hello people! I hope you like ProtoTool-Launcher. I spent quite a while making this to make managing and launching companion tools as easy and seamless as possible. Since this is a solo project, some bugs are expected so please report any issues or feedback in my Discord server!

### [**The Discord Server**](https://discord.gg/YDXpQPF6g9)


---


##  Installation & Usage

### Windows
1. Go to the **Releases** tab on GitHub
2. Download the latest `ProtoTool-Launcher.exe`
3. Run the executable and enjoy!

*(I have made this unbelievably simple, I believe you can do it! :3)*

### Linux (Tested & Supported)
To run the ToolBox directly from source on Linux:
```bash
# Clone the repository
git clone https://github.com/CaptainBoots/ProtoTool-Launcher.git
cd ProtoTool-Launcher

# Run with Python
python3 ProtoTool-Launcher.py
```

*(If you encounter any platform-specific issues on Linux, please let me know in the Discord server!)*

---

## Self-Updater

Under the hood, ProtoTool-Launcher features a custom-engineered self-update mechanism designed specifically for Windows PyInstaller packages:
* **No-Lock Rename:** Because Windows locks active running executables, the ToolBox dynamically renames its running binary to `ToolBox.exe.bak` (supporting incrementing suffixes like `.bak.1`, `.bak.2` to resolve any local collisions).
* **Safe Downloading:** It streams the new compiled binary from GitHub Releases directly to the original path. If the network drops or the file is unavailable, it automatically rolls back your previous executable so you never lose your app.
* **Automated Cleanup:** On next launch, the application uses wildcard pattern globbing to identify and cleanly purge all old `.bak` backup files from your directory.

---

##  Troubleshooting 

### Antivirus False Positives
PyInstaller executables are sometimes flagged as false positives by Windows Defender or other security scanners. If the executable is blocked or cannot be deleted/compiled, add the installation or workspace folder to your antivirus exclusion list.

### Locked Executables
If compiling with PyInstaller fails with a `PermissionError` (Access Denied), ensure that PyCharm, VS Code, or any background instances of `ToolBox.exe` are completely closed so they release their locks.

---

i hope you all like the priject and if you wana make a tool join the discord above.

*Made with <3 by:*
1. Boots @CaptainBoots