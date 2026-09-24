@echo off
setlocal
title Allison's Data Extractor - Build

cd /d "%~dp0"

echo ============================================
echo    Allison's Data Extractor Build System
echo ============================================
echo.

echo [1/7] Checking Python...
py --version
if errorlevel 1 (
    echo.
    echo ERROR: Python could not be found.
    pause
    exit /b 1
)

echo.
echo [2/7] Checking PyInstaller...
py -m PyInstaller --version
if errorlevel 1 (
    echo.
    echo PyInstaller is not installed. Installing...
    py -m pip install pyinstaller
    if errorlevel 1 goto :failed
)

echo.
echo [3/7] Checking source dependencies...
py -c "import pdfplumber, openpyxl, PIL, pytesseract"
if errorlevel 1 (
    echo Installing required Python packages...
    py -m pip install pdfplumber openpyxl pillow pytesseract
    if errorlevel 1 goto :failed
)

echo.
echo [4/7] Removing previous build output...
if exist "build" rmdir /s /q "build"
if exist "build_updater" rmdir /s /q "build_updater"
if exist "dist" rmdir /s /q "dist"
if exist "dist_updater" rmdir /s /q "dist_updater"
if exist "release" rmdir /s /q "release"
if exist "Allisons Data Extractor Updater.spec" del /q "Allisons Data Extractor Updater.spec"

echo.
echo [5/7] Building main portable application...
py -m PyInstaller --noconfirm --clean "Allisons Data Extractor.spec"
if errorlevel 1 goto :failed

echo.
echo [6/7] Building updater helper...
py -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Allisons Data Extractor Updater" ^
  --distpath "dist_updater" ^
  --workpath "build_updater" ^
  updater.py
if errorlevel 1 goto :failed

copy /Y "dist_updater\Allisons Data Extractor Updater.exe" "dist\Allisons Data Extractor\Allisons Data Extractor Updater.exe" >nul
if errorlevel 1 goto :failed

echo.
echo [7/7] Creating USB update package...
py package_update.py
if errorlevel 1 goto :failed

echo.
echo ============================================
echo BUILD SUCCESSFUL
echo ============================================
echo.
echo Portable application:
echo   dist\Allisons Data Extractor\
echo.
echo Future USB update package:
echo   release\Allisons_Data_Extractor_Update_v*.zip
echo.
echo Copy the entire dist\Allisons Data Extractor folder for a full install.
echo After revision 1.2.0 is installed, future ZIP updates can be loaded from
 echo Settings ^> Backend / Administrator ^> Load Update Package.
echo.
pause
exit /b 0

:failed
echo.
echo ============================================
echo BUILD FAILED
echo ============================================
echo Review the error messages above.
echo.
pause
exit /b 1
