@echo off
title Telly Handel Import - Build Executable
color 0a

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║              TELLY HANDEL IMPORT BUILDER                 ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

echo 🏗️  Building your Tally Excel Processor into an executable...
echo 🐍 Detected Python 3.13 - using simple build approach...
echo.

python build_simple.py

echo.
echo ✅ Build process completed!
echo.
echo 📦 Your executable should be in the 'dist' folder
echo 📂 Look for: TellyHandelImport.exe
echo.
echo Press any key to open the dist folder...
pause >nul

if exist "dist\TellyHandelImport.exe" (
    echo Opening dist folder...
    explorer dist
) else (
    echo ❌ Executable not found. Check for errors above.
)

pause