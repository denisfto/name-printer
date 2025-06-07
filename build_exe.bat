@echo off
echo =============================================
echo    WINDOWS NAME PRINTER - BUILD SCRIPT
echo =============================================
echo.

echo Installing required Python packages...
pip install pywin32 Pillow pyinstaller

echo.
echo Building executable file...
pyinstaller --onefile --windowed --name="NamePrinter" --icon=printer.ico windows_printer_app.py

echo.
echo =============================================
echo Build completed!
echo.
echo The executable file is located in the 'dist' folder:
echo   dist\NamePrinter.exe
echo.
echo You can now distribute this single .exe file
echo to any Windows computer (no Python required).
echo =============================================
echo.
pause