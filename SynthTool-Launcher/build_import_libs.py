import os
import subprocess
import sys

PYSIDE6_DIR = r"C:\Users\boots\AppData\Local\Programs\Python\Python314\Lib\site-packages\PySide6"
MSVC_BIN_DIR = r"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\MSVC\14.44.35207\bin\Hostx64\x64"

DUMPBIN_PATH = os.path.join(MSVC_BIN_DIR, "dumpbin.exe")
LIB_PATH = os.path.join(MSVC_BIN_DIR, "lib.exe")

DLL_NAMES = ["Qt6Core.dll", "Qt6Gui.dll", "Qt6Widgets.dll", "Qt6Network.dll"]
OUTPUT_LIB_DIR = r"C:\Users\boots\PycharmProjects\VRChat-ToolBox\SynthTool-Launcher\lib"

os.makedirs(OUTPUT_LIB_DIR, exist_ok=True)

print(f"PYSIDE6_DIR: {PYSIDE6_DIR}")
print(f"MSVC_BIN_DIR: {MSVC_BIN_DIR}")
print(f"OUTPUT_LIB_DIR: {OUTPUT_LIB_DIR}")

for dll_name in DLL_NAMES:
    dll_path = os.path.join(PYSIDE6_DIR, dll_name)
    lib_name = dll_name.replace(".dll", ".lib")
    def_name = dll_name.replace(".dll", ".def")
    
    def_path = os.path.join(OUTPUT_LIB_DIR, def_name)
    out_lib_path = os.path.join(OUTPUT_LIB_DIR, lib_name)
    
    print(f"\nProcessing {dll_name}...")
    
    # 1. Run dumpbin /exports
    try:
        cmd = [DUMPBIN_PATH, "/exports", dll_path]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running dumpbin for {dll_name}: {e.stderr}")
        continue
        
    # 2. Parse exported symbols
    symbols = []
    lines = result.stdout.splitlines()
    start_parsing = False
    
    for line in lines:
        line_strip = line.strip()
        if not start_parsing:
            if "ordinal hint RVA      name" in line_strip:
                start_parsing = True
            continue
            
        if not line_strip:
            if len(symbols) > 0:
                # Empty line after symbols list means end of exports section
                break
            continue
            
        parts = line_strip.split()
        if len(parts) >= 4:
            rva = parts[2]
            # Verify RVA is a valid 8-digit hex offset
            if len(rva) == 8 and all(c in "0123456789ABCDEFabcdef" for c in rva):
                symbol_name = parts[3]
                # Filter out forwarded symbols or junk symbols if any
                if symbol_name.startswith("[") or "(" in symbol_name:
                    continue
                symbols.append(symbol_name)
        elif len(parts) == 3:
            # Sometime hint is missing, parts is [ordinal, RVA, name]
            rva = parts[1]
            if len(rva) == 8 and all(c in "0123456789ABCDEFabcdef" for c in rva):
                symbol_name = parts[2]
                if symbol_name.startswith("[") or "(" in symbol_name:
                    continue
                symbols.append(symbol_name)

    print(f"Discovered {len(symbols)} exported symbols.")
    
    # 3. Write DEF file
    dll_base_name = dll_name.replace(".dll", "")
    with open(def_path, "w", encoding="utf-8") as f:
        f.write(f"LIBRARY {dll_base_name}\n")
        f.write("EXPORTS\n")
        for sym in symbols:
            f.write(f"    {sym}\n")
            
    print(f"Wrote DEF file: {def_path}")
    
    # 4. Generate LIB file using lib.exe
    try:
        cmd = [LIB_PATH, f"/def:{def_path}", f"/out:{out_lib_path}", "/machine:x64"]
        print(f"Running: {' '.join(cmd)}")
        lib_result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
        print(f"Successfully generated LIB file: {out_lib_path}")
        
        # Clean up temporary def file and exp file
        os.remove(def_path)
        exp_path = out_lib_path.replace(".lib", ".exp")
        if os.path.exists(exp_path):
            os.remove(exp_path)
            
    except subprocess.CalledProcessError as e:
        print(f"Error generating LIB for {dll_name}: {e.stderr}")
        print(f"Stdout: {e.stdout}")
