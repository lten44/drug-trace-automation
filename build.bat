@echo off
title Build Installer

echo ============================================
echo  Build Installer for Drug Trace App
echo ============================================
echo.

cd /d "E:\Workbuddy\2026-05-29-13-47-22\drug_trace"

echo [Step 1/3] PyInstaller building...
pyinstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo [ERROR] PyInstaller failed!
    pause
    exit /b 1
)
echo [OK] PyInstaller done.
echo.

echo [Step 2/3] Inno Setup compiling...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup failed!
    pause
    exit /b 1
)
echo [OK] Installer created.
echo.

echo [Step 3/3] Opening output folder...
start .

echo.
echo ============================================
echo  Build complete!
echo  Check output folder for Setup-*.exe
echo ============================================
echo.
pause
