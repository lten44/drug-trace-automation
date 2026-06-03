@echo off
chcp 65001 >nul
title 药品追溯码 v3.0 一键发布
cd /d "D:\workplace\drug_trace_app"

echo ============================================
echo  药品追溯码 v3.0 - 一键发布
echo ============================================
echo.

:: ── 1. 推送到 GitHub ──
echo [1/4] 推送到 GitHub...
git add -A
git commit -m "v3.0 统一排重模式+输入表匹配+UI美化"
git push origin main
if %errorlevel% neq 0 (
    echo.
    echo ⚠️ 推送失败，可能是未登录 GitHub
    echo 请手动执行: git push origin main
    pause
    exit /b 1
)
echo [OK] 推送成功！
echo.

:: ── 2. 清理旧构建 ──
echo [2/4] 清理旧构建文件...
rmdir /s /q dist 2>nul
rmdir /s /q build 2>nul
echo [OK] 清理完成
echo.

:: ── 3. PyInstaller 打包 ──
echo [3/4] PyInstaller 打包 exe...
pyinstaller "药品批发企业追朔码自动处理软件.spec" --noconfirm --clean
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller 失败！
    pause
    exit /b 1
)
echo [OK] exe 打包完成
echo.

:: ── 4. Inno Setup 制作安装包 ──
echo [4/4] 制作安装包...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if %errorlevel% neq 0 (
    echo [ERROR] Inno Setup 失败！请确认已安装 Inno Setup 6
    pause
    exit /b 1
)
echo [OK] 安装包制作完成！
echo.

:: ── 找到生成的安装包 ──
for %%f in (Setup-*.exe) do set SETUP_FILE=%%f
echo.
echo ============================================
echo  ✅ 全部完成！
echo.
echo  安装包：%SETUP_FILE%
echo.
echo  下一步：打开 GitHub 创建 Release
echo  https://github.com/lten44/drug-trace-automation/releases/new
echo.
echo  Tag: v3.0
echo  标题: 药品批发企业追朔码自动处理软件 v3.0
echo  附件: 拖拽 %SETUP_FILE% 到附件区
echo ============================================
echo.
start .
pause