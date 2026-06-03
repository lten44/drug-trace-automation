param([switch]$SkipPush)

$root = "D:\workplace\drug_trace_app"
Set-Location $root

Write-Host "===== Drug Trace App v3.0 Publish =====" -ForegroundColor Cyan

# 1. Git push
if (-not $SkipPush) {
    Write-Host "[1/4] Pushing to GitHub..." -ForegroundColor Yellow
    git add -A
    git commit -m "v3.0 release"
    git push origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[!] Push failed. Run: git push origin main" -ForegroundColor Red
        Write-Host "    Or use: publish.ps1 -SkipPush" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "[OK] Push done" -ForegroundColor Green
} else {
    Write-Host "[1/4] Skip push" -ForegroundColor Yellow
}

# 2. Clean
Write-Host "[2/4] Cleaning old builds..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
Write-Host "[OK] Cleaned" -ForegroundColor Green

# 3. PyInstaller
Write-Host "[3/4] Building exe..." -ForegroundColor Yellow
$spec = Get-Item "*.spec" | Where-Object { $_.Name.StartsWith("药品") }
pyinstaller $spec.FullName --noconfirm --clean
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] PyInstaller failed" -ForegroundColor Red
    exit 1
}
Write-Host "[OK] exe built" -ForegroundColor Green

# 4. Inno Setup
Write-Host "[4/4] Building installer..." -ForegroundColor Yellow
$iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (Test-Path $iscc) {
    & $iscc "installer.iss"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Inno Setup failed" -ForegroundColor Red
        exit 1
    }
    Write-Host "[OK] Installer built" -ForegroundColor Green
} else {
    Write-Host "[!] Inno Setup not found at:" -ForegroundColor Red
    Write-Host "    $iscc" -ForegroundColor Red
    Write-Host "    exe is in dist/ folder, install package skipped" -ForegroundColor Yellow
}

# Done
$setup = Get-ChildItem "Setup-*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
Write-Host ""
Write-Host "===== DONE =====" -ForegroundColor Green
if ($setup) {
    Write-Host "Installer: $($setup.Name)" -ForegroundColor Green
} else {
    Write-Host "EXE path: dist/ ?" -ForegroundColor Green
}
Write-Host ""
Write-Host "Next: Create GitHub Release" -ForegroundColor Yellow
Write-Host "https://github.com/lten44/drug-trace-automation/releases/new" -ForegroundColor Yellow
Write-Host "Tag: v3.0  |  Asset: drag the Setup-*.exe" -ForegroundColor White

Start-Process "."
pause