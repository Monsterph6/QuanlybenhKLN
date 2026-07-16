# -*- coding: utf-8 -*-
# Cap nhat QuanLyBenhNhanTHA len ban moi nhat tu GitHub Releases (repo private).
# Khong dung Python/git tren may dich - chi can PowerShell (co san tren Windows).
# Script nay nam CUNG thu muc voi QuanLyBenhNhanTHA.exe va tu cap nhat tai cho.
# Neu may nay duoc cai voi vai tro "May chu" (co san QuanLyBenhNhanTHA-Service.exe),
# script se tu dong dung dich vu chia se truoc khi thay file, roi bat lai sau
# khi xong - buoc nay can quyen Administrator.

$ErrorActionPreference = "Stop"

$GITHUB_OWNER = "Monsterph6"
$GITHUB_REPO  = "quanlybenhnhantha"

$root            = $PSScriptRoot
$versionFile     = Join-Path $root "VERSION.txt"
$tokenFile       = Join-Path $root "update_token.txt"
$exePath         = Join-Path $root "QuanLyBenhNhanTHA.exe"
$internalDir     = Join-Path $root "_internal"
$serviceExe      = Join-Path $root "QuanLyBenhNhanTHA-Service.exe"
$serviceInternal = Join-Path $root "_internal_service"
$trayExe         = Join-Path $root "QuanLyBenhNhanTHA-Tray.exe"
$trayInternal    = Join-Path $root "_internal_tray"
$isServerRole    = Test-Path $serviceExe

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn2($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Err2($msg)  { Write-Host $msg -ForegroundColor Red }

# --- Neu may nay la May chu, can quyen Administrator de dung/bat lai dich vu ---
if ($isServerRole) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Err2 "May nay dang o vai tro May chu (co Windows Service dang chia se du lieu) - can chay file nay voi quyen Administrator (chuot phai update.bat, chon 'Run as administrator')."
        exit 1
    }
}

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

if ($isServerRole) {
    $trayRunning = Get-Process -Name "QuanLyBenhNhanTHA-Tray" -ErrorAction SilentlyContinue
    if ($trayRunning) {
        Write-Warn2 "Bảng điều khiển Máy chủ (tray) đang chạy. Vui lòng chuột phải icon khay hệ thống, chọn 'Thoát' trước khi cập nhật."
        Read-Host "Đóng tray xong, nhấn Enter để tiếp tục (hoặc Ctrl+C để hủy)"
    }
}

# --- Lay token (chi can neu repo dang o che do Private; bo trong duoc neu
# repo da chuyen sang Public - xem README.md muc "Lay Personal Access Token") ---
if (-not (Test-Path $tokenFile)) {
    Write-Warn2 "Chua cau hinh Personal Access Token."
    Write-Host "Neu repo dang Private: xem huong dan lay token trong README.md, muc 'Lay Personal Access Token'."
    Write-Host "Neu repo da chuyen sang Public: cu de trong roi nhan Enter, khong can token."
    $token = Read-Host "Dan Personal Access Token vao day (de trong neu khong can)"
    Set-Content -Path $tokenFile -Value $token.Trim() -NoNewline -Encoding utf8
}
$token = (Get-Content $tokenFile -Raw).Trim()
$headers = @{
    Accept       = "application/vnd.github+json"
    "User-Agent" = "QuanLyBenhNhanTHA-Updater"
}
if ($token) {
    $headers["Authorization"] = "token $token"
}

# --- Phien ban hien tai ---
$localVersion = "0.0.0"
if (Test-Path $versionFile) {
    $localVersion = (Get-Content $versionFile -Raw).Trim()
}
Write-Info "Phien ban hien tai: $localVersion"

# --- Hoi GitHub xem ban moi nhat la gi ---
$releasesUrl = "https://api.github.com/repos/$GITHUB_OWNER/$GITHUB_REPO/releases?per_page=20"
try {
    $releases = Invoke-RestMethod -Uri $releasesUrl -Headers $headers
} catch {
    Write-Err2 "Khong the ket noi GitHub hoac token khong hop le / khong co quyen truy cap repo."
    Write-Err2 "Chi tiet loi: $($_.Exception.Message)"
    Write-Host "Neu token sai, xoa file update_token.txt roi chay lai de nhap token moi."
    exit 1
}
$release = $releases | Where-Object { $_.tag_name -match '^v\d' -and -not $_.draft -and -not $_.prerelease } | Select-Object -First 1
if (-not $release) {
    Write-Err2 "Khong tim thay Release nao (tag dang 'vX.Y.Z') trong repo."
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
    Accept       = "application/octet-stream"
    "User-Agent" = "QuanLyBenhNhanTHA-Updater"
}
if ($token) {
    $downloadHeaders["Authorization"] = "token $token"
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

if ($isServerRole) {
    Write-Info "Dang dung dich vu chia se truoc khi cap nhat..."
    & $serviceExe stop
    Start-Sleep -Seconds 2
}

Write-Info "Dang cai dat ban moi (giu nguyen benh_nhan.db, lan_config.json, backups\\) ..."
if (Test-Path $internalDir) { Remove-Item $internalDir -Recurse -Force }
if (Test-Path $exePath) { Remove-Item $exePath -Force }
Move-Item -Path $newInternal -Destination $internalDir -Force
Move-Item -Path $newExe -Destination $exePath -Force

if ($isServerRole) {
    $newServiceExe = Join-Path $tmpExtract "QuanLyBenhNhanTHA-Service.exe"
    $newServiceInternal = Join-Path $tmpExtract "_internal_service"
    $newTrayExe = Join-Path $tmpExtract "QuanLyBenhNhanTHA-Tray.exe"
    $newTrayInternal = Join-Path $tmpExtract "_internal_tray"
    if (Test-Path $newServiceExe) {
        if (Test-Path $serviceInternal) { Remove-Item $serviceInternal -Recurse -Force }
        if (Test-Path $serviceExe) { Remove-Item $serviceExe -Force }
        Move-Item -Path $newServiceInternal -Destination $serviceInternal -Force
        Move-Item -Path $newServiceExe -Destination $serviceExe -Force
    }
    if (Test-Path $newTrayExe) {
        if (Test-Path $trayInternal) { Remove-Item $trayInternal -Recurse -Force }
        if (Test-Path $trayExe) { Remove-Item $trayExe -Force }
        Move-Item -Path $newTrayInternal -Destination $trayInternal -Force
        Move-Item -Path $newTrayExe -Destination $trayExe -Force
    }
}

Set-Content -Path $versionFile -Value $remoteVersion -NoNewline -Encoding utf8

if ($isServerRole) {
    Write-Info "Dang bat lai dich vu chia se..."
    & $serviceExe start
}

Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
Remove-Item $tmpExtract -Recurse -Force -ErrorAction SilentlyContinue

Write-Ok "Da cap nhat thanh cong len phien ban $remoteVersion."
Write-Ok "Chay lai QuanLyBenhNhanTHA.exe de su dung."
