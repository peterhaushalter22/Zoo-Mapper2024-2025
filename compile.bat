@echo off
echo Building Zoo Mapper Executable...

:: Navigate to the root directory (where this script is located)
cd /d %~dp0

:: Clean previous builds (optional)
rmdir /s /q build
rmdir /s /q dist
del /q *.spec

:: Run PyInstaller with all resources
::pyinstaller --noconfirm --onefile --windowed src\main\zoo.py ^
::--add-data "src\\main\\resources\\icons;src\\main\\resources\\icons" ^
::--add-data "src\\main\\resources\\Logo.jpg;src\\main\\resources\\Logo.jpg"
pyinstaller --onefile --windowed ^
--icon src\main\resources\clienticon.ico ^
--add-data "src\\main\\resources\\icons;src/main/resources/icons" ^
--add-data "src\\main\\resources\\Logo.jpg;src/main/resources/Logo.jpg" ^
src\main\zoo.py


echo.
echo Build Complete! Executable is located in the dist folder.
pause
