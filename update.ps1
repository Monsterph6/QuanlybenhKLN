# -*- coding: utf-8 -*-
# Cap nhat QuanLyBenhNhanTHA len ban moi nhat tu GitHub Releases (repo private).
# Khong dung Python/git tren may dich - chi can PowerShell (co san tren Windows).
# Script nay nam CUNG thu muc voi QuanLyBenhNhanTHA.exe va tu cap nhat tai cho.

$ErrorActionPreference = "Stop"

$GITHUB_OWNER = "Monsterph6"
$GITHUB_REPO  = "quanlybenhnhantha"

$root        = $PSScriptRoot
$versionFile = Join-Path $root "VERSION.txt"
$tokenFile   = Join-Path $root "update_token.txt"
$exePath     = Join-Path $root "QuanLyBenhNhanTHA.exe"
$internalDir = Join-Path $root "_internal"

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Err2($msg)  { Write-Host $msg -ForegroundColor Red }

# --- Dam bao ung dung dang khong chay (khong the ghi de file .exe/.dll dang mo) ---
$running = Get-Process -Name "QuanLyBenhNhanTHA" -ErrorAction SilentlyContinue
if ($running) {
    Write-Warn2 "Ung dung dang chay. Vui long dong QuanLyBenhNhanTHA.exe truoc khi cap nhat."
    Read-Host "Dong ung dung xong, nhan Enter de tiep tuc (hoac Ctrl+C de huy)"
    $running = Get-Process -Name "QuanLyBenhNhanTHA" -ErrorAction SilentlyContinue
    if ($running) {
        Write-Err2 "Ung dung van dang chay. Huy cap nhat."
        exit 1
    }
}

# --- Lay token (repo private can Personal Access Token de tai duoc) ---
if (-not (Test-Path $tokenFile)) {
    Write-Warn2 "Chua co token GitHub de tai ban cap nhat (repo o che do private)."
    Write-Host "Xem huong dan lay token trong README.md, muc 'Lay Personal Access Token'."
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

# --- Tim file .zip dinh kem trong Release (ban portable, khong phai Setup.exe) ---
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

$newExe = Join-Path $tmpExtract "QuanLyBenhNhanTHA.exe"
$newInternal = Join-Path $tmpExtract "_internal"
if (-not (Test-Path $newExe) -or -not (Test-Path $newInternal)) {
    Write-Err2 "File tai ve khong dung dinh dang mong doi (thieu QuanLyBenhNhanTHA.exe hoac _internal)."
    exit 1
}

Write-Info "Dang cai dat ban moi (giu nguyen du lieu benh_nhan.db) ..."
if (Test-Path $internalDir) { Remove-Item $internalDir -Recurse -Force }
if (Test-Path $exePath) { Remove-Item $exePath -Force }
Move-Item -Path $newInternal -Destination $internalDir -Force
Move-Item -Path $newExe -Destination $exePath -Force

Set-Content -Path $versionFile -Value $remoteVersion -NoNewline -Encoding utf8

Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
Remove-Item $tmpExtract -Recurse -Force -ErrorAction SilentlyContinue

Write-Ok "Da cap nhat thanh cong len phien ban $remoteVersion."
Write-Ok "Chay lai QuanLyBenhNhanTHA.exe de su dung."
