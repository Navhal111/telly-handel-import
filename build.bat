@echo off
echo ========================================
echo    Telly Handel Import - Build Script
echo ========================================
echo.

echo [1/5] Checking Python installation...
python --version
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    pause
    exit /b 1
)

echo.
echo [2/5] Installing core packages individually (Python 3.13 compatible)...
echo Installing pandas...
pip install pandas --upgrade
echo Installing openpyxl...
pip install openpyxl --upgrade
echo Installing xlrd...
pip install xlrd --upgrade
echo Installing requests...
pip install requests --upgrade
echo Installing pillow...
pip install pillow --upgrade

echo.
echo [3/5] Installing PyInstaller...
pip install pyinstaller --upgrade
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller!
    pause
    exit /b 1
)

echo.
echo [4/5] Cleaning previous builds...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo [5/5] Building executable...
echo This may take a few minutes...
pyinstaller build.spec --clean --noconfirm
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo           BUILD SUCCESSFUL!
echo ========================================
echo.
echo The executable has been created:
echo Location: dist\TellyHandelImport.exe
echo.
echo You can now distribute this .exe file to other computers
echo without needing Python installed!
echo.
if exist "dist\TellyHandelImport.exe" (
    dir dist\TellyHandelImport.exe
) else (
    echo Warning: Executable not found in expected location
    dir dist\*.exe
)
echo.
pause