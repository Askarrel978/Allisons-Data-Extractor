@echo off
title Allison's PDF Extractor - Build

echo ============================================
echo    Allison's PDF Extractor Build System
echo ============================================
echo.

cd /d "%~dp0"

echo [1/4] Checking Python...
py --version
if errorlevel 1 (
    echo.
    echo ERROR: Python could not be found.
    pause
    exit /b 1
)

echo.
echo [2/4] Checking PyInstaller...
py -m PyInstaller --version
if errorlevel 1 (
    echo.
    echo PyInstaller is not installed.
    echo Installing PyInstaller...
    py -m pip install pyinstaller
)

echo.
echo [3/4] Removing previous build...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo [4/4] Building Allison's PDF Extractor...
py -m PyInstaller --noconfirm --clean "Allisons PDF Extractor.spec"

if errorlevel 1 (
    echo.
    echo ============================================
    echo BUILD FAILED
    echo ============================================
    pause
    exit /b 1
)

echo.
echo ============================================
echo BUILD SUCCESSFUL
echo ============================================
echo.
echo Your application is located at:
echo.
echo dist\Allisons PDF Extractor\
echo.
pause