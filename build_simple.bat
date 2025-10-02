@echo off
echo Building Cargos App for Windows (Simple Method)...

REM Install dependencies
pip install -r requirements.txt

REM Clean previous builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM Build executable with all local modules explicitly included
pyinstaller --onefile --windowed --name "CargosApp" ^
    --add-data "templates;templates" ^
    --add-data "config.json;." ^
    --hidden-import models ^
    --hidden-import services ^
    --hidden-import ui_components ^
    --hidden-import config_manager ^
    --hidden-import unified_config_service ^
    --hidden-import validators ^
    --hidden-import constants ^
    main.py

REM Create distribution package
mkdir dist_windows
copy dist\CargosApp.exe dist_windows\
copy templates dist_windows\templates
copy config.json dist_windows\

echo.
echo Build complete! Check dist_windows folder.
echo Run CargosApp.exe to test.
pause
