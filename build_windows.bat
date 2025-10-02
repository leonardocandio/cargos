@echo off
echo Building Cargos App for Windows...

REM Install dependencies
pip install -r requirements.txt
pip install pyinstaller

REM Clean previous builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM Build executable using spec file
pyinstaller --clean CargosApp.spec

REM Create distribution package
mkdir dist_windows
copy dist\CargosApp\* dist_windows\
copy templates dist_windows\templates
copy config.json dist_windows\

REM Create README for Windows users
echo # Cargos App - Windows > dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Installation >> dist_windows\README.txt
echo 1. Extract all files to a folder on your Windows computer >> dist_windows\README.txt
echo 2. Run CargosApp.exe >> dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Requirements >> dist_windows\README.txt
echo - Windows 10 or later >> dist_windows\README.txt
echo - No additional software required (all dependencies included) >> dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Usage >> dist_windows\README.txt
echo 1. Click "Browse" to select your Excel file >> dist_windows\README.txt
echo 2. Select templates (CARGO and/or AUTORIZACION) >> dist_windows\README.txt
echo 3. Click "Generate Documents" to create Word documents >> dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Configuration >> dist_windows\README.txt
echo - The config.json file contains all settings and can be edited >> dist_windows\README.txt
echo - Changes to config.json will be preserved between app restarts >> dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Troubleshooting >> dist_windows\README.txt
echo - If the app doesn't start, try running as Administrator >> dist_windows\README.txt
echo - Make sure your Excel file is in the correct format >> dist_windows\README.txt
echo - Check that template files are in the templates/ folder >> dist_windows\README.txt
echo. >> dist_windows\README.txt
echo ## Support >> dist_windows\README.txt
echo For issues or questions, contact the development team. >> dist_windows\README.txt

echo.
echo Build complete! Check dist_windows folder.
echo The config.json file is included and will be writable.
pause
