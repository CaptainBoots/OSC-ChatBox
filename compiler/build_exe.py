import subprocess
import sys
import os

# Change to the compiler directory
os.chdir(r'C:\Users\boots\PycharmProjects\OSC-ChatBox\compiler')

print("=" * 60)
print("OSC-Installer EXE Builder")
print("=" * 60)

# Step 1: Install PyInstaller
print("\n[1/3] Installing PyInstaller...")
try:
    result = subprocess.run(
        [sys.executable, '-m', 'pip', 'install', 'pyinstaller', '-q'],
        check=True,
        capture_output=True,
        text=True
    )
    print("✓ PyInstaller installed successfully")
except subprocess.CalledProcessError as e:
    print(f"✗ Failed to install PyInstaller: {e.stderr}")
    sys.exit(1)

# Step 2: Build the EXE
print("\n[2/3] Building EXE with PyInstaller...")
try:
    result = subprocess.run(
        [
            sys.executable, '-m', 'PyInstaller',
            '--onefile',
            '--windowed',
            '--name', 'VRChat-ToolBox',
            '--hidden-import=requests',
            'VRChat-ToolBox.py'
        ],
        check=True,
        capture_output=True,
        text=True
    )
    print("✓ EXE built successfully")
except subprocess.CalledProcessError as e:
    print(f"✗ Build failed: {e.stderr}")
    sys.exit(1)

# Step 3: Verify
print("\n[3/3] Verifying output...")
exe_path = r'C:\Users\boots\PycharmProjects\OSC-ChatBox\compiler\dist\VRChat-ToolBox.exe'
if os.path.exists(exe_path):
    size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"✓ EXE created successfully!")
    print(f"\n📍 Location: {exe_path}")
    print(f"📦 Size: {size_mb:.2f} MB")
else:
    print(f"✗ EXE file not found at {exe_path}")
    sys.exit(1)

print("\n" + "=" * 60)
print("Build complete! You can now run VRChat-ToolBox.exe")
print("=" * 60)
