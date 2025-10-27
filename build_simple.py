#!/usr/bin/env python3
"""
Simple build script for Python 3.13 - installs packages individually
"""

import subprocess
import sys
import os

def install_package(package):
    """Install a single package."""
    print(f"📦 Installing {package}...")
    try:
        result = subprocess.run([
            sys.executable, "-m", "pip", "install", package, "--upgrade"
        ], capture_output=True, text=True, check=True)
        print(f"✅ {package} installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install {package}")
        print(f"   Error: {e.stderr}")
        return False

def main():
    print("🏗️  SIMPLE BUILD FOR PYTHON 3.13")
    print("="*50)
    
    # Required packages
    packages = [
        "pandas",
        "openpyxl", 
        "xlrd",
        "requests",
        "pillow",
        "pyinstaller"
    ]
    
    # Install packages
    failed_packages = []
    for package in packages:
        if not install_package(package):
            failed_packages.append(package)
    
    if failed_packages:
        print(f"\n⚠️  Some packages failed to install: {failed_packages}")
        print("   Continuing with build anyway...")
    
    # Clean previous builds
    print("\n🧹 Cleaning previous builds...")
    if os.path.exists("build"):
        import shutil
        shutil.rmtree("build")
    if os.path.exists("dist"):
        import shutil
        shutil.rmtree("dist")
    
    # Build executable
    print("\n🔨 Building executable...")
    try:
        result = subprocess.run([
            "pyinstaller", "build.spec", "--clean", "--noconfirm"
        ], check=True)
        
        # Check result
        exe_path = "dist/TellyHandelImport.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024*1024)
            print(f"\n🎉 BUILD SUCCESSFUL!")
            print(f"📦 Executable: {exe_path}")
            print(f"📏 Size: {size_mb:.1f} MB")
            print(f"\n✅ Ready to distribute!")
            return True
        else:
            print(f"\n❌ Build completed but executable not found")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Build failed: {e}")
        return False

if __name__ == "__main__":
    success = main()
    print(f"\n{'='*50}")
    if success:
        print("🎯 All done! Your .exe is ready in the dist/ folder")
    else:
        print("❌ Build failed - check errors above")
    
    input("\nPress Enter to exit...")
