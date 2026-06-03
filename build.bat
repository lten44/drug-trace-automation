@echo off
title Build v3.0 Installer

echo ============================================
echo  Build v3.0 Installer for Drug Trace App
echo ============================================
echo.

cd /d "D:\workplace\drug_trace_app"

echo [Step 1/3] Cleaning old builds...
rmdir /s /q dist 2>nul
rmdir /s /q build 2>nul
echo [OK] Cleaned.

echo [Step 2/3] PyInstaller building...
pyinstaller "build-v3.spec" --noconfirm --clean
if errorlevel 1 (
    echo [ERROR] PyInstaller failed!
    pause
    exit /b 1
)
echo [OK] PyInstaller done.
echo.

echo [Step 3/3] Inno Setup compiling...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup failed!
    pause
    exit /b 1
)
echo [OK] Installer created: Setup-*.exe
echo.

echo [Done] Opening output folder...
start .

echo.
echo ============================================
echo  Build complete!
echo  Output: Setup-药品批发企业追朔码自动处理软件-v3.0.exe
echo ============================================
echo.
pause