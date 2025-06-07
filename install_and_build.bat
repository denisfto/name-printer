@echo off
echo =============================================
echo    NAME PRINTER - INSTALLATION & BUILD
echo =============================================
echo.

echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.7+ from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python found!
python --version

echo.
echo Installing required packages...
echo Installing pywin32...
pip install pywin32
echo Installing Pillow...
pip install Pillow
echo Installing PyInstaller...
pip install pyinstaller

echo.
echo Building executable...
if not exist "windows_printer_app.py" (
    echo ERROR: windows_printer_app.py not found!
    echo Make sure all files are in the same folder.
    pause
    exit /b 1
)

pyinstaller --onefile --windowed --name="NamePrinter" windows_printer_app.py

echo.
if exist "dist\NamePrinter.exe" (
    echo =============================================
    echo SUCCESS! Build completed successfully!
    echo =============================================
    echo.
    echo The executable file has been created:
    echo   %CD%\dist\NamePrinter.exe
    echo.
    echo You can now:
    echo 1. Copy NamePrinter.exe to any Windows computer
    echo 2. Run it without installing Python
    echo 3. Select your printer on first run
    echo 4. Start printing names!
    echo.
    echo Opening the dist folder...
    explorer dist
) else (
    echo =============================================
    echo BUILD FAILED!
    echo =============================================
    echo Check the error messages above.
)

echo.
pause