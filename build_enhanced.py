#!/usr/bin/env python3
"""
Enhanced build script for creating the Telly Handel Import executable.
This script ensures all dependencies are properly bundled.
"""

import os
import sys
import subprocess
import shutil

def run_command(command, description):
    """Run a command and handle errors."""
    print(f"\n{'='*50}")
    print(f"🔄 {description}")
    print(f"{'='*50}")
    
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} - SUCCESS")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} - FAILED")
        print(f"Error: {e}")
        if e.stdout:
            print(f"Output: {e.stdout}")
        if e.stderr:
            print(f"Error Output: {e.stderr}")
        return False

def check_python():
    """Check if Python is properly installed."""
    try:
        result = subprocess.run([sys.executable, "--version"], capture_output=True, text=True)
        print(f"✅ Python version: {result.stdout.strip()}")
        return True
    except:
        print("❌ Python not found!")
        return False

def main():
    """Main build process."""
    print("🏗️  TELLY HANDEL IMPORT - BUILD SCRIPT")
    print("="*60)
    
    # Check Python
    if not check_python():
        return False
    
    # Install PyInstaller
    if not run_command(f"{sys.executable} -m pip install pyinstaller", "Installing PyInstaller"):
        return False
    
    # Install requirements - try Python 3.13 specific first, then fallback
    python_version = sys.version_info
    print(f"🐍 Detected Python {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version >= (3, 13):
        print("📦 Using Python 3.13+ compatible requirements...")
        if not run_command(f"{sys.executable} -m pip install -r requirements_py313.txt", "Installing Python 3.13 dependencies"):
            print("⚠️  Python 3.13 requirements failed, trying individual packages...")
            # Try installing packages individually without strict version constraints
            packages = ["pandas", "openpyxl", "xlrd", "requests", "pillow", "pyinstaller"]
            for package in packages:
                if not run_command(f"{sys.executable} -m pip install {package}", f"Installing {package}"):
                    print(f"⚠️  Failed to install {package}, continuing...")
    else:
        if not run_command(f"{sys.executable} -m pip install -r requirements.txt", "Installing dependencies"):
            return False
    
    # Clean previous builds
    if os.path.exists("build"):
        shutil.rmtree("build")
        print("🧹 Cleaned build directory")
    
    if os.path.exists("dist"):
        shutil.rmtree("dist")
        print("🧹 Cleaned dist directory")
    
    # Build executable
    if not run_command("pyinstaller build.spec --clean --noconfirm", "Building executable"):
        return False
    
    # Check if executable was created
    exe_path = "dist/ExcelProcessor.exe"
    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path) / (1024*1024)  # MB
        print("\n" + "="*60)
        print("🎉 BUILD SUCCESSFUL!")
        print("="*60)
        print(f"📦 Executable created: {exe_path}")
        print(f"📏 File size: {file_size:.1f} MB")
        print("\n📋 DISTRIBUTION NOTES:")
        print("   ✅ This .exe file is standalone")
        print("   ✅ No Python installation required on target machine")
        print("   ✅ All dependencies are bundled")
        print("   ✅ Ready for distribution!")
        print("\n🚀 You can now share ExcelProcessor.exe with others!")
        return True
    else:
        print("❌ Build failed - executable not found")
        return False

if __name__ == "__main__":
    success = main()
    input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)