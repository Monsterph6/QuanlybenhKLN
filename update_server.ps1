# -*- coding: utf-8 -*-
# Cap nhat goi MAY CHU (QuanLyBenhNhanTHA-Service + QuanLyBenhNhanTHA-Tray)
# len ban moi nhat tu GitHub Releases (repo private, tag "server-vX.Y.Z" -
# tach rieng dong phien ban voi ung dung may tram, dung "update.ps1").
# CAN CHAY VOI QUYEN ADMINISTRATOR (de dung/bat lai Windows Service).
# Script nay nam CUNG thu muc voi QuanLyBenhNhanTHA-Service.exe va tu cap
# nhat tai cho.

$ErrorActionPreference = "Stop"

$GITHUB_OWNER = "Monsterph6"
$GITHUB_REPO  = "quanlybenhnhantha"

$root           = $PSScriptRoot
$versionFile    = Join-Path $root "VERSION_SERVER.txt"
$tokenFile      = Join-Path $root "update_token.txt"
$serviceExe     = Join-Path $root "QuanLyBenhNhanTHA-Service.exe"
$serviceInternal = Join-Path $root "_internal_service"
$trayExe        = Join-Path $root "QuanLyBenhNhanTHA-Tray.exe"
$trayInternal   = Join-Path $root "_internal_tray"

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Err2($msg)  { Write-Host $msg -ForegroundColor Red }

# --- Can quyen Administrator de dung/bat Windows Service ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Err2 "Can chay file nay voi quyen Administrator (chuot phai update_server.bat, chon 'Run as administrator')."
    exit 1
}

if (-not (Test-Path $serviceExe)) {
    Write-Err2 "Khong tim thay QuanLyBenhNhanTHA-Service.exe cung thu muc voi script nay."
    Write-Err2 "Script nay dung cho ban da dong goi (.exe) - neu chay tu ma nguon, dung 'git pull' de cap nhat."
    exit 1
}

# --- Dong tray helper truoc neu dang chay (khoa file .exe dang mo) ---
$trayRunning = Get-Process -Name "QuanLyBenhNhanTHA-Tray" -ErrorAction SilentlyContinue
if ($trayRunning) {
    Write-Warn2 "Tray helper dang chay. Vui long chuot phai icon khay he thong, chon 'Thoat' truoc khi cap nhat."
    Read-Host "Dong tray xong, nhan Enter de tiep tuc (hoac Ctrl+C de huy)"
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
    "User-Agent"  = "QuanLyBenhNhanTHA-Server-Updater"
}

# --- Phien ban hien tai ---
$localVersion = "0.0.0"
if (Test-Path $versionFile) {
    $localVersion = (Get-Content $versionFile -Raw).Trim()
}
Write-Info "Phien ban may chu hien tai: $localVersion"

# --- Hoi GitHub xem ban moi nhat la gi (chi tag "server-vX.Y.Z") ---
$releasesUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/releases?per_page=20"
try {
    $releases = Invoke-RestMethod -Uri $releasesUrl -Headers $headers
} catch {
    Write-Err2 "Khong the ket noi GitHub hoac token khong hop le / khong co quyen truy cap repo."
    Write-Err2 "Chi tiet loi: $($_.Exception.Message)"
    Write-Host "Neu token sai, xoa file update_token.txt roi chay lai de nhap token moi."
    exit 1
}
$release = $releases | Where-Object { $_.tag_name -match '^server-v\d' -and -not $_.draft -and -not $_.prerelease } | Select-Object -First 1
if (-not $release) {
    Write-Err2 "Khong tim thay Release nao cho may chu (tag dang 'server-vX.Y.Z') trong repo."
    exit 1
}

$remoteVersion = $release.tag_name -replace '^server-v', ''
Write-Info "Phien ban may chu moi nhat tren GitHub: $remoteVersion"

if ($remoteVersion -eq $localVersion) {
    Write-Ok "Ban dang dung la ban moi nhat roi. Khong can cap nhat."
    exit 0
}

# --- Tim file .zip dinh kem trong Release ---
$asset = $release.assets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1
if (-not $asset) {
    Write-Err2 "Khong tim thay file dinh kem (.zip) trong Release '$($release.tag_name)'."
    exit 1
}

Write-Info "Dang tai ban $remoteVersion ($([math]::Round($asset.size/1MB,1)) MB) ..."
$downloadHeaders = @{
    Authorization = "token $token"
    Accept        = "application/octet-stream"
    "User-Agent"  = "QuanLyBenhNhanTHA-Server-Updater"
}
$tmpZip = Join-Path $env:TEMP "QuanLyBenhNhanTHA-Server-update.zip"
Invoke-WebRequest -Uri $asset.url -Headers $downloadHeaders -OutFile $tmpZip

$tmpExtract = Join-Path $env:TEMP "QuanLyBenhNhanTHA-Server-update-extract"
if (Test-Path $tmpExtract) { Remove-Item $tmpExtract -Recurse -Force }
Expand-Archive -Path $tmpZip -DestinationPath $tmpExtract -Force

$newServiceExe = Join-Path $tmpExtract "QuanLyBenhNhanTHA-Service.exe"
$newServiceInternal = Join-Path $tmpExtract "_internal_service"
$newTrayExe = Join-Path $tmpExtract "QuanLyBenhNhanTHA-Tray.exe"
$newTrayInternal = Join-Path $tmpExtract "_internal_tray"
if (-not (Test-Path $newServiceExe) -or -not (Test-Path $newServiceInternal)) {
    Write-Err2 "File tai ve khong dung dinh dang mong doi (thieu QuanLyBenhNhanTHA-Service.exe hoac _internal_service)."
    exit 1
}

# --- Dung dich vu truoc khi thay file (khong the ghi de .exe/.dll dang chay) ---
Write-Info "Dang dung dich vu chia se..."
& $serviceExe stop
Start-Sleep -Seconds 2

Write-Info "Dang cai dat ban moi (giu nguyen benh_nhan.db, lan_config.json, backups\\) ..."
if (Test-Path $serviceInternal) { Remove-Item $serviceInternal -Recurse -Force }
if (Test-Path $serviceExe) { Remove-Item $serviceExe -Force }
Move-Item -Path $newServiceInternal -Destination $serviceInternal -Force
Move-Item -Path $newServiceExe -Destination $serviceExe -Force

if (Test-Path $newTrayExe) {
    if (Test-Path $trayInternal) { Remove-Item $trayInternal -Recurse -Force }
    if (Test-Path $trayExe) { Remove-Item $trayExe -Force }
    Move-Item -Path $newTrayInternal -Destination $trayInternal -Force
    Move-Item -Path $newTrayExe -Destination $trayExe -Force
}

Set-Content -Path $versionFile -Value $remoteVersion -NoNewline -Encoding utf8

Write-Info "Dang bat lai dich vu..."
& $serviceExe start

Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
Remove-Item $tmpExtract -Recurse -Force -ErrorAction SilentlyContinue

Write-Ok "Da cap nhat may chu thanh cong len phien ban $remoteVersion."
Write-Ok "Dich vu QuanLyBenhNhanTHA_Server da duoc bat lai. Mo lai server_tray.py / QuanLyBenhNhanTHA-Tray.exe neu can xem trang thai."
