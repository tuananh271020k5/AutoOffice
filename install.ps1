<#
.SYNOPSIS
    Cài đặt AutoOffice Add-in cho Microsoft Word, Excel và PowerPoint thông qua Trusted Catalog.
.DESCRIPTION
    Script này cấu hình môi trường Microsoft Office theo hướng dẫn chính thức của Microsoft:
    1. Tạo thư mục chia sẻ cục bộ C:\AutoOfficeAddin
    2. Sao chép/Tải manifest.production.xml vào thư mục
    3. Đăng ký thư mục mạng vào Registry Trusted Catalogs của Office
    Hoàn toàn bằng mã nguồn mở PowerShell minh bạch, không sử dụng file nhị phân thực thi (.exe).
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

# Kiểm tra quyền Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[!] Script can quyen Administrator de tao thu muc chia se mang." -ForegroundColor Yellow
    Write-Host "[*] Dang khoi chay lai voi quyen Administrator..." -ForegroundColor Cyan
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "         BAT DAU CAI DAT AUTOOFFICE ADD-IN                " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

$ShareName = "AutoOfficeAddin"
$LocalPath = "C:\AutoOfficeAddin"
$CatalogGuid = "{B2C3D4E5-F6A7-8901-BCDE-F12345678903}"
$ComputerName = $env:COMPUTERNAME
if (-not $ComputerName) { $ComputerName = "localhost" }
$NetworkPath = "\\$ComputerName\$ShareName"
$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"

# 1. Dong cac tien trinh Office dang mo de tranh xung dot
Write-Host "[1/5] Kiem tra va dong cac tien trinh Office..." -ForegroundColor Cyan
Get-Process -Name "WINWORD", "EXCEL", "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# 2. Tao thu muc cuc bo
Write-Host "[2/5] Tao thu muc cai dat tai: $LocalPath" -ForegroundColor Cyan
if (-not (Test-Path $LocalPath)) {
    New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
}

# 3. Sao chep manifest.xml vao thu muc
Write-Host "[3/5] Thiet lap file manifest..." -ForegroundColor Cyan
$ManifestSource = Join-Path $PSScriptRoot "manifest.production.xml"
$ManifestDest = Join-Path $LocalPath "manifest.xml"

if (Test-Path $ManifestSource) {
    Copy-Item -Path $ManifestSource -Destination $ManifestDest -Force
} else {
    $ManifestUrl = "https://raw.githubusercontent.com/tuananh271020k5/AutoOffice/master/manifest.production.xml"
    Write-Host "     Dang tai manifest tu: $ManifestUrl" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $ManifestUrl -OutFile $ManifestDest -UseBasicParsing
}

# 4. Tao Network Share cuc bo cho Office
Write-Host "[4/5] Cau hinh Network Share noi bo ($ShareName)..." -ForegroundColor Cyan
& net share $ShareName /delete /y 2>$null | Out-Null
& net share "$ShareName=$LocalPath" "/GRANT:Everyone,FULL" | Out-Null

# 5. Dang ky vao Registry Trusted Catalogs
Write-Host "[5/5] Dang ky Trusted Catalog vao Microsoft Office..." -ForegroundColor Cyan
if (-not (Test-Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs")) {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" -Force | Out-Null
}
if (-not (Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force | Out-Null
}

Set-ItemProperty -Path $RegistryPath -Name "Id" -Value $CatalogGuid -Type String
Set-ItemProperty -Path $RegistryPath -Name "Url" -Value $NetworkPath -Type String
Set-ItemProperty -Path $RegistryPath -Name "Flags" -Value 1 -Type DWord

# Xoa cache WEF cu de Office nhan Add-in ngay lap tuc
$WefCache = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
if (Test-Path $WefCache) {
    Remove-Item -Path $WefCache -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "           CAI DAT AUTOOFFICE THANH CONG!                " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "Huong dan mo tren Microsoft Word:" -ForegroundColor Yellow
Write-Host "1. Mo Microsoft Word."
Write-Host "2. Vao the 'Home' -> Chon 'Add-ins' (hoac Insert -> Get Add-ins)."
Write-Host "3. Chon tab 'SHARED FOLDER' (Thu muc dung chung) -> Chon AutoOffice -> Bam 'Add'."
Write-Host ""
Read-Host "Nhan Enter de ket thuc..."
