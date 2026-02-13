"""
Build Script for Chat Status Monitor
=====================================
Creates a standalone Windows executable (.exe)

Usage:
    python build_exe.py

Requirements:
    pip install pyinstaller
"""

import subprocess
import sys
import os

def main():
    print("=" * 60)
    print("  Building Chat Status Monitor Standalone Executable")
    print("=" * 60)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✓ PyInstaller found")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed")
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=ChatStatusMonitor",
        "--onefile",              # Single .exe file
        "--windowed",             # No console window (GUI app)
        "--noconfirm",            # Overwrite without asking
        "--clean",                # Clean build files
        "--add-data", f"README.md{os.pathsep}.",  # Include README
        "chat_monitor_gui.py"
    ]
    
    print("\nRunning PyInstaller...")
    print(f"Command: {' '.join(cmd)}\n")
    
    try:
        subprocess.check_call(cmd)
        print("\n" + "=" * 60)
        print("  BUILD SUCCESSFUL!")
        print("=" * 60)
        print("\nYour executable is at:")
        print("  dist/ChatStatusMonitor.exe")
        print("\nTo use it:")
        print("  1. Copy ChatStatusMonitor.exe to any folder")
        print("  2. Make sure Tesseract OCR is installed")
        print("  3. Double-click to run!")
        print("\nNote: First run may take a few seconds to start.")
        
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error: {e}")
        print("\nTry running manually:")
        print("  pip install pyinstaller")
        print("  pyinstaller --onefile --windowed chat_monitor_gui.py")

if __name__ == "__main__":
    main()