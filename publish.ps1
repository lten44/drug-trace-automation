param(
    [switch]$SkipPush = $false
)

$ErrorActionPreference = "Stop"
$root = "D:\workplace\drug_trace_app"
Set-Location $root

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  药品追溯码 v3.0 - 一键发布" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# 1. Git push
if (-not $SkipPush) {
    Write-Host "[1/4] 推送到 GitHub..." -ForegroundColor Yellow
    git add -A
    git commit -m "v3.0 统一排重模式+输入表匹配+UI美化"
    git push origin main 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "`n[!] 推送失败，请手动执行: git push origin main" -ForegroundColor Red
        Write-Host "    或用 -SkipPush 跳过推送直接打包" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "[OK] 推送成功！" -ForegroundColor Green
} else {
    Write-Host "[1/4] 跳过推送" -ForegroundColor Yellow
}
Write-Host ""

# 2. Clean
Write-Host "[2/4] 清理旧构建..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
Write-Host "[OK] 清理完成" -ForegroundColor Green
Write-Host ""

# 3. PyInstaller
Write-Host "[3/4] PyInstaller 打包 exe..." -ForegroundColor Yellow
$spec = Get-Item "药品批发企业追朔码自动处理软件.spec"
pyinstaller $spec.FullName --noconfirm --clean
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] PyInstaller 失败" -ForegroundColor Red
    exit 1
}
Write-Host "[OK] exe 打包完成" -ForegroundColor Green
Write-Host ""

# 4. Inno Setup
Write-Host "[4/4] 制作安装包..." -ForegroundColor Yellow
$iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (Test-Path $iscc) {
    & $iscc "installer.iss"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Inno Setup 失败" -ForegroundColor Red
        exit 1
    }
    Write-Host "[OK] 安装包制作完成" -ForegroundColor Green
} else {
    Write-Host "[!] 未找到 Inno Setup: $iscc" -ForegroundColor Red
    Write-Host "    跳过安装包生成，dist 目录中已有 exe" -ForegroundColor Yellow
}
Write-Host ""

# 查找生成的安装包
$setup = Get-ChildItem "Setup-*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  [完成]" -ForegroundColor Green
if ($setup) {
    Write-Host "  安装包: $($setup.Name)" -ForegroundColor Green
} else {
    Write-Host "  exe 路径: dist\药品批发企业追朔码自动处理软件\" -ForegroundColor Green
}
Write-Host ""
Write-Host "  下一步：创建 GitHub Release" -ForegroundColor Yellow
Write-Host "  https://github.com/lten44/drug-trace-automation/releases/new" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Tag: v3.0" -ForegroundColor White
Write-Host "  标题: 药品批发企业追朔码自动处理软件 v3.0" -ForegroundColor White
Write-Host "  说明: 打开 RELEASE_NOTE_v3.0.md 复制" -ForegroundColor White
Write-Host "  附件: 拖拽安装包到附件区" -ForegroundColor White
Write-Host "============================================" -ForegroundColor Cyan

Start-Process "."
pause