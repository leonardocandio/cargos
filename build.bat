@echo off
echo Building File Generator App...

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

REM Build executable
echo Building executable with PyInstaller...
pyinstaller --onefile --windowed --name="FileGenerator" --add-data "templates;templates" --hidden-import="config_manager" --hidden-import="validators" --hidden-import="constants" main.py

echo Build complete! Executable is in dist/ folder
pause
