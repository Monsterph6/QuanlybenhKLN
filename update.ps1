# -*- coding: utf-8 -*-
# Cap nhat QuanLyBenhNhanTHA len ban moi nhat tu GitHub Releases (repo private).
# Khong dung Python/git tren may dich - chi can PowerShell (co san tren Windows).

$ErrorActionPreference = "Stop"

$GITHUB_OWNER = "phamhieu111089"
$GITHUB_REPO  = "quanlybenhnhantha"

$root        = $PSScriptRoot
$versionFile = Join-Path $root "VERSION.txt"
$tokenFile   = Join-Path $root "update_token.txt"
$exeDir      = Join-Path $root "QuanLyBenhNhanTHA"
$exePath     = Join-Path $exeDir "QuanLyBenhNhanTHA.exe"

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Err2($msg)  { Write-Host $msg -ForegroundColor Red }

# --- Lay token (repo private can Personal Access Token de tai duoc) ---
if (-not (Test-Path $tokenFile)) {
    Write-Warn2 "Chua co token GitHub de tai ban cap nhat (repo o che do private)."
    Write-Host "Xem huong dan lay token trong README_MAY_DICH.md, muc 'Lay Personal Access Token'."
    $token = Read-Host "Dan Personal Access Token vao day roi Enter (token se duoc luu lai cho lan sau)"
    if ([string]::IsNullOrWhiteSpace($token)) {
        Write-Err2 "Chua nhap token. Huy cap nhat."
        exit 1
    }
    Set-Content -Path $tokenFile -Value $token.Trim() -NoNewline -Encoding utf8
}
$token = (Get-Content $tokenFile -Raw).Trim()
$headers = @{
    Authorization = "token $token"
    Accept        = "application/vnd.github+json"
    "User-Agent"  = "QuanLyBenhNhanTHA-Updater"
}

# --- Phien ban hien tai ---
$localVersion = "0.0.0"
if (Test-Path $versionFile) {
    $localVersion = (Get-Content $versionFile -Raw).Trim()
}
Write-Info "Phien ban hien tai: $localVersion"

# --- Hoi GitHub xem ban moi nhat la gi ---
$releaseUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/releases/latest"
try {
    $release = Invoke-RestMethod -Uri $releaseUrl -Headers $headers
} catch {
    Write-Err2 "Khong the ket noi GitHub hoac token khong hop le / khong co quyen truy cap repo."
    Write-Err2 "Chi tiet loi: $($_.Exception.Message)"
    Write-Host "Neu token sai, xoa file update_token.txt roi chay lai de nhap token moi."
    exit 1
}

$remoteVersion = $release.tag_name -replace '^v', ''
Write-Info "Phien ban moi nhat tren GitHub: $remoteVersion"

if ($remoteVersion -eq $localVersion) {
    Write-Ok "Ban dang dung la ban moi nhat roi. Khong can cap nhat."
    exit 0
}

# --- Tim file .exe (zip) dinh kem trong Release ---
$asset = $release.assets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1
if (-not $asset) {
    Write-Err2 "Khong tim thay file dinh kem (.zip) trong Release '$($release.tag_name)'."
    exit 1
}

Write-Info "Dang tai ban $remoteVersion ($([math]::Round($asset.size/1MB,1)) MB) ..."
$downloadHeaders = @{
    Authorization = "token $token"
    Accept        = "application/octet-stream"
    "User-Agent"  = "QuanLyBenhNhanTHA-Updater"
}
$tmpZip = Join-Path $env:TEMP "QuanLyBenhNhanTHA-update.zip"
Invoke-WebRequest -Uri $asset.url -Headers $downloadHeaders -OutFile $tmpZip

$tmpExtract = Join-Path $env:TEMP "QuanLyBenhNhanTHA-update-extract"
if (Test-Path $tmpExtract) { Remove-Item $tmpExtract -Recurse -Force }
Expand-Archive -Path $tmpZip -DestinationPath $tmpExtract -Force

# Zip co the giai nen ra 1 thu muc con (vd QuanLyBenhNhanTHA/) hoac ngay tai goc
$srcDir = $tmpExtract
$inner = Join-Path $tmpExtract "QuanLyBenhNhanTHA"
if (Test-Path $inner) { $srcDir = $inner }

Write-Info "Dang cai dat ban moi (giu nguyen du lieu benh_nhan.db) ..."
if (Test-Path $exeDir) {
    Remove-Item $exeDir -Recurse -Force
}
Move-Item -Path $srcDir -Destination $exeDir -Force

Set-Content -Path $versionFile -Value $remoteVersion -NoNewline -Encoding utf8

Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
Remove-Item $tmpExtract -Recurse -Force -ErrorAction SilentlyContinue

Write-Ok "Da cap nhat thanh cong len phien ban $remoteVersion."
Write-Ok "Chay lai QuanLyBenhNhanTHA.exe (hoac bam Run.bat) de su dung."
