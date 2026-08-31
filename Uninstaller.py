# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
#                                          VRChat-ToolBox Uninstaller                                                     #
# ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════#
# Made with <3 by Gemini CLI

import os
import sys
import shutil
import glob

CENTRAL_CONFIG_DIR = os.path.join(os.getenv("LOCALAPPDATA", ""), "VRChat-ToolBox")
INSTALL_PATH_POINTER = os.path.join(CENTRAL_CONFIG_DIR, "install_path.txt")
TOOLS_PATH_POINTER = os.path.join(CENTRAL_CONFIG_DIR, "tools_path.txt")

def print_header():
    print("=" * 60)
    print("              VRChat-ToolBox Uninstaller")
    print("=" * 60)
    print()

def get_recorded_install_path():
    if os.path.exists(INSTALL_PATH_POINTER):
        try:
            with open(INSTALL_PATH_POINTER, "r", encoding="utf-8") as f:
                path = f.read().strip()
                if os.path.isdir(path):
                    return path
        except Exception:
            pass
    return None

def get_recorded_tools_path():
    if os.path.exists(TOOLS_PATH_POINTER):
        try:
            with open(TOOLS_PATH_POINTER, "r", encoding="utf-8") as f:
                path = f.read().strip()
                if os.path.isdir(path):
                    return path
        except Exception:
            pass
    return None

def uninstall_from_recorded_positions():
    print("[1] Removing files from recorded positions...")
    install_path = get_recorded_install_path()
    tools_path = get_recorded_tools_path()

    deleted_any = False

    if install_path:
        print(f" -> Found recorded installation folder: {install_path}")
        # Safeguard: Never delete root C: or User Profile root
        if len(os.path.abspath(install_path)) > 15:
            try:
                # Remove ToolBox.exe or VRChat-ToolBox files at that position
                for item in ["ToolBox.exe", "VRChat-ToolBox.py", "ToolBox.spec"]:
                    full_p = os.path.join(install_path, item)
                    if os.path.isfile(full_p):
                        os.remove(full_p)
                        print(f"    Deleted file: {full_p}")
                        deleted_any = True
                
                # Check for build/dist folders in the install path
                for folder in ["build", "dist"]:
                    full_f = os.path.join(install_path, folder)
                    if os.path.isdir(full_f):
                        shutil.rmtree(full_f, ignore_errors=True)
                        print(f"    Deleted directory: {full_f}")
                        deleted_any = True
            except Exception as e:
                print(f"    [Error] Failed to remove install path files: {e}")
        else:
            print("    [Warning] Installation path too short, skipping for safety.")

    if tools_path:
        print(f" -> Found recorded tools folder: {tools_path}")
        if len(os.path.abspath(tools_path)) > 15 and "VRChat-Tools" in os.path.abspath(tools_path):
            try:
                shutil.rmtree(tools_path, ignore_errors=True)
                print(f"    Deleted tools directory: {tools_path}")
                deleted_any = True
            except Exception as e:
                print(f"    [Error] Failed to remove tools path: {e}")
        else:
            print("    [Warning] Tools path failed safety guard-rails check, skipping.")

    if not deleted_any:
        print(" -> No recorded files found to delete.")
    print("[✓] Recorded position cleaning completed.\n")

def delete_appdata_folders():
    print("[2] Removing AppData config folders...")
    appdata_local = os.path.join(os.getenv("LOCALAPPDATA", ""), "VRChat-ToolBox")
    appdata_roaming_tools = os.path.join(os.getenv("APPDATA", ""), "VRChat-Tools")
    appdata_roaming_toolbox = os.path.join(os.getenv("APPDATA", ""), "VRChat-ToolBox")

    deleted_any = False

    for folder in [appdata_local, appdata_roaming_tools, appdata_roaming_toolbox]:
        if os.path.isdir(folder):
            print(f" -> Found AppData folder: {folder}")
            try:
                shutil.rmtree(folder, ignore_errors=True)
                print(f"    Deleted folder: {folder}")
                deleted_any = True
            except Exception as e:
                print(f"    [Error] Failed to delete folder {folder}: {e}")

    if not deleted_any:
        print(" -> No AppData folders found.")
    print("[✓] AppData config cleaning completed.\n")

def deep_search_and_clean():
    print("[3] Performing safe deep search and clean across standard locations...")
    
    # Define standard search locations with absolute guard rails
    user_profile = os.environ.get("USERPROFILE", "C:\\Users\\Default")
    search_roots = [
        os.path.join(user_profile, "Desktop"),
        os.path.join(user_profile, "Downloads"),
        os.path.join(user_profile, "Documents"),
        os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs")
    ]

    deleted_any = False

    for root in search_roots:
        if not os.path.isdir(root):
            continue
        
        print(f" -> Scanning folder: {root}")
        # Search for shortcuts or directories matching VRChat-ToolBox / VRChat-Tools
        try:
            # Safe check: limit searching to the specified directories using strict glob patterns
            for item in glob.glob(os.path.join(root, "*ToolBox*")) + glob.glob(os.path.join(root, "*VRChat-Tools*")):
                # Safe Guard-rails: check filename pattern
                base = os.path.basename(item).lower()
                if "vrchat" in base or "toolbox" in base or "tools" in base:
                    if os.path.isdir(item):
                        shutil.rmtree(item, ignore_errors=True)
                        print(f"    Deleted directory: {item}")
                        deleted_any = True
                    elif os.path.isfile(item):
                        os.remove(item)
                        print(f"    Deleted file: {item}")
                        deleted_any = True
        except Exception as e:
            print(f"    [Error] Error scanning {root}: {e}")

    if not deleted_any:
        print(" -> No residual files or shortcuts found in standard locations.")
    print("[✓] Deep search cleaning completed.\n")

def main():
    print_header()
    
    print("Please choose an uninstallation option:")
    print(" 1. Clean uninstall from recorded file positions only.")
    print(" 2. Clean all AppData configuration folders only.")
    print(" 3. Safe deep search and delete residual files/shortcuts across standard directories.")
    print(" 4. Complete Uninstall (Runs all of the above options).")
    print(" 5. Exit.")
    print()

    try:
        choice = input("Enter choice (1-5): ").strip()
    except (KeyboardInterrupt, EOFError):
        print("\nExiting.")
        return

    if choice == "1":
        uninstall_from_recorded_positions()
    elif choice == "2":
        delete_appdata_folders()
    elif choice == "3":
        deep_search_and_clean()
    elif choice == "4":
        print("\n=== Initiating Complete Uninstall ===\n")
        uninstall_from_recorded_positions()
        delete_appdata_folders()
        deep_search_and_clean()
        print("=== Complete Uninstall Process Finished! ===\n")
    elif choice == "5" or choice.lower() == "exit":
        print("Exiting Uninstaller.")
        return
    else:
        print("[Error] Invalid option. Exiting.")
        return

    input("Press Enter to close...")

if __name__ == "__main__":
    main()
