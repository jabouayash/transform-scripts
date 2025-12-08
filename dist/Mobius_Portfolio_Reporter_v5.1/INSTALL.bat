@echo off
echo ==========================================
echo  Mobius Portfolio Reporter - Installer
echo  Version 5.1
echo ==========================================
echo.

:: Create folder structure
echo Creating folders...
mkdir "C:\Mobius Reports" 2>nul
mkdir "C:\Mobius Reports\Incoming" 2>nul
mkdir "C:\Mobius Reports\Transformed" 2>nul
mkdir "C:\Mobius Reports\Archive" 2>nul

:: Copy Excel file
echo Copying Portfolio Transformer...
copy /Y "%~dp0files\Portfolio Transformer.xlsm" "C:\Mobius Reports\"

echo.
echo ==========================================
echo  Folders created and files copied!
echo ==========================================
echo.
echo NEXT STEPS:
echo 1. Open docs\SETUP_GUIDE.md for detailed instructions
echo 2. Set up the Outlook email monitor (paste code from files\OutlookMonitor.txt)
echo 3. Enable Bloomberg Excel Add-in
echo.
echo Press any key to open the setup guide...
pause >nul
start "" "%~dp0docs\SETUP_GUIDE.md"
