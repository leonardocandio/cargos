@echo off
REM Windows build script for File Generator App
REM This script creates a standalone executable for Windows

echo Building File Generator App for Windows...

REM Create virtual environment if it doesn't exist
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install requirements
echo Installing requirements...
pip install -r requirements.txt

REM Build executable with cross-platform data handling
echo Building executable with PyInstaller...
pyinstaller --onefile --windowed --name="FileGenerator" --add-data "templates;templates" --hidden-import="config_manager" --hidden-import="validators" --hidden-import="constants" --hidden-import="docxtpl" --hidden-import="docxcompose" main.py

echo Build complete! Executable is in dist/ folder
echo.
echo Note: For cross-platform builds, use build.sh on Unix systems
pause
