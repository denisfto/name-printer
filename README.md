# Name Printer for Windows

A simple Windows application to print names in large uppercase letters on A4 landscape format.

## Features

- **Automatic font sizing**: Maximizes font size based on the longest text
- **Printer selection**: Choose from available system printers (saved for future use)
- **Uppercase conversion**: Automatically converts all text to uppercase
- **A4 landscape layout**: Optimized for horizontal printing
- **Additional text support**: Secondary text 15% smaller than main text
- **Direct printing**: Prints directly to selected printer
- **No installation required**: Single .exe file

## Quick Start

### Option 1: Use Pre-built Executable
1. Download `NamePrinter.exe`
2. Double-click to run
3. Select your printer on first use
4. Enter names and print!

### Option 2: Build from Source
1. Extract all files to a folder
2. Double-click `install_and_build.bat`
3. Wait for build to complete
4. Find `NamePrinter.exe` in the `dist` folder

## Requirements for Building

- Windows 7/8/10/11
- Python 3.7 or newer
- Internet connection (for downloading packages)

## How to Use

### First Time Setup
1. Run `NamePrinter.exe`
2. A printer selection dialog will appear
3. Choose your printer from the list
4. Click "Select Printer"
5. Settings are automatically saved

### Printing Names
1. Enter **First Name** (optional)
2. Enter **Last Name** (optional)
3. Add **Additional Text** if needed (optional)
4. Click "Update Preview" to see how it will look
5. Click "**PRINT NOW**" to print

### Example Output
```
JOHN SMITH              ← Large font (auto-sized)
KONSTANTINOPOL DIMITROV ← Same large font
                        
ADDITIONAL INFORMATION  ← 15% smaller font
CONTACT DETAILS         ← 15% smaller font
```

## Features in Detail

### Automatic Font Sizing
- Program finds the longest text (first name or last name)
- Calculates maximum font size to fit ~90% of page width
- Both names use the same large font size
- Additional text uses 85% of main font size

### Printer Management
- First run: Shows printer selection dialog
- Saves your choice in `printer_settings.json`
- Change printer: Click "Setup Printer" button
- Supports both local and network printers

### Text Processing
- All text automatically converted to UPPERCASE
- Left-aligned layout
- Smart spacing between sections
- A4 landscape orientation (297×210mm)

## File Structure

```
NamePrinter/
├── windows_printer_app.py    # Main application code
├── requirements.txt          # Python dependencies
├── install_and_build.bat    # Automatic build script
├── build_exe.bat           # Simple build script
├── README.md               # This file
└── dist/
    └── NamePrinter.exe     # Built executable (after build)
```

## Troubleshooting

### Build Issues
- **Python not found**: Install Python from [python.org](https://python.org) and check "Add to PATH"
- **Package installation fails**: Run as Administrator
- **Build fails**: Check if all files are in the same folder

### Printing Issues
- **No printers found**: Check if printers are installed in Windows
- **Print fails**: Try "Setup Printer" and select a different printer
- **Wrong orientation**: Make sure printer supports A4 landscape
- **Size issues**: Check printer settings for "Fit to page" or "Actual size"

### Runtime Issues
- **Won't start**: Install Visual C++ Redistributable 2019+
- **Slow startup**: Normal on first run, faster afterwards
- **Settings lost**: Check if `printer_settings.json` exists in same folder

## Technical Details

- **Language**: Python 3.7+ with tkinter GUI
- **Dependencies**: pywin32, Pillow, PyInstaller
- **Print resolution**: 300 DPI for quality output
- **Image format**: PNG temporary files for printing
- **Font**: Uses system Arial font (fallback to default if unavailable)

## Support

For issues or questions:
1. Check this README file
2. Verify printer is working with other applications
3. Try rebuilding with latest code
4. Check Windows Event Viewer for detailed errors

## Version History

- **v1.0**: Initial release with basic printing functionality
- Font auto-sizing, printer selection, uppercase conversion
- A4 landscape layout, additional text support

---
