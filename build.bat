@echo off
echo Installing Python dependencies...
pip install -r requirements.txt

echo.
echo Building executable...
pyinstaller build.spec

echo.
echo Build complete! The executable can be found in the 'dist' folder.
echo.
pause