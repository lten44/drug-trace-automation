param(
    [Parameter(Mandatory=$false)]
    [string]$Token = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Help
)

if ($Help) {
    Write-Host "Usage: .\create-release.ps1 -Token YOUR_GITHUB_TOKEN"
    Write-Host ""
    Write-Host "How to get a token:"
    Write-Host "  1. Go to https://github.com/settings/tokens"
    Write-Host "  2. Click 'Generate new token (classic)'"
    Write-Host "  3. Select scope: repo"
    Write-Host "  4. Copy the token and pass it here"
    exit
}

if (-not $Token) {
    Write-Host "[!] No token provided" -ForegroundColor Red
    Write-Host "    Usage: .\create-release.ps1 -Token YOUR_TOKEN" -ForegroundColor Yellow
    Write-Host "    Or: .\create-release.ps1 -Help" -ForegroundColor Yellow
    exit 1
}

$owner = "lten44"
$repo = "drug-trace-automation"
$tag = "v3.0"
$name = "药品批发企业追朔码自动处理软件 v3.0"
$body = Get-Content "RELEASE_NOTE_v3.0.md" -Raw
$assetPath = "Setup-药品批发企业追朔码自动处理软件-v3.0.exe"
$headers = @{
    "Authorization" = "token $Token"
    "Accept" = "application/vnd.github.v3+json"
}

Write-Host "===== Creating GitHub Release v3.0 =====" -ForegroundColor Cyan
Write-Host ""

# 1. Create release
Write-Host "[1/3] Creating release..." -ForegroundColor Yellow
$releaseBody = @{
    tag_name = $tag
    name = $name
    body = $body
    draft = $false
    prerelease = $false
} | ConvertTo-Json

try {
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases" `
        -Method Post -Headers $headers -Body $releaseBody -ContentType "application/json"
    Write-Host "[OK] Release created: $($release.html_url)" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to create release: $_" -ForegroundColor Red
    exit 1
}

# 2. Upload asset
Write-Host "[2/3] Uploading installer..." -ForegroundColor Yellow
$uploadUrl = $release.upload_url -replace "\{.*\}", ""
$assetHeaders = @{
    "Authorization" = "token $Token"
    "Accept" = "application/vnd.github.v3+json"
    "Content-Type" = "application/x-msdownload"
}

try {
    $asset = Invoke-RestMethod -Uri "$uploadUrl?name=$assetPath" `
        -Method Post -Headers $assetHeaders -InFile $assetPath
    Write-Host "[OK] Installer uploaded: $($asset.browser_download_url)" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to upload: $_" -ForegroundColor Red
    Write-Host "    Release was created but asset upload failed. Upload manually." -ForegroundColor Yellow
}

# 3. Done
Write-Host ""
Write-Host "===== DONE =====" -ForegroundColor Green
Write-Host "Release URL: https://github.com/$owner/$repo/releases/tag/$tag" -ForegroundColor Cyan
Write-Host ""
pause