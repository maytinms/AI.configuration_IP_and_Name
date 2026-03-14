# Set-NetworkConfig.ps1
param([string]$CsvPath = ".\config.csv")
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CsvFile = Join-Path $ScriptDir $CsvPath

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "MAC自动配置工具 v1.0" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $CsvFile)) {
    Write-Host "[错误] 找不到配置文件: $CsvFile" -ForegroundColor Red; exit 1
}

Write-Host "[1] 读取配置文件..." -ForegroundColor Yellow
$csvData = Import-Csv -Path $CsvFile

Write-Host "[2] 获取本机网卡MAC地址..." -ForegroundColor Yellow
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.MacAddress -ne $null -and $_.InterfaceDescription -notmatch "Virtual|Loopback|Bluetooth|VMware|VirtualBox|Hyper-V" }

if ($adapters.Count -eq 0) { Write-Host "[错误] 未找到有效的网卡" -ForegroundColor Red; exit 1 }

Write-Host "     发现网卡:" -ForegroundColor Gray
foreach ($adapter in $adapters) { Write-Host "       - $($adapter.Name): $($adapter.MacAddress -replace '-','')" -ForegroundColor Gray }

$targetConfig = $null
foreach ($adapter in $adapters) {
    $mac = $adapter.MacAddress -replace '-', ''
    foreach ($row in $csvData) {
        if ($row.Mac.Trim().ToLower() -eq $mac.ToLower()) {
            $targetConfig = $row; $matchedAdapter = $adapter; break
        }
    }
    if ($targetConfig) { break }
}

if (-not $targetConfig) { Write-Host "[错误] CSV中未找到匹配的MAC地址" -ForegroundColor Red; exit 1 }

Write-Host "[匹配成功] 计算机名: $($targetConfig.Name) IP: $($targetConfig.IP)" -ForegroundColor Green
Write-Host "[3] 设置网络配置..." -ForegroundColor Yellow

$interfaceAlias = $matchedAdapter.Name
$ipAddress = $targetConfig.IP
$gateway = $targetConfig.Gateway
$dns1 = $targetConfig.DNS1; $dns2 = $targetConfig.DNS2

# 子网掩码转前缀长度
$parts = $targetConfig.Sub -split '\.'
$count = 0
foreach ($p in $parts) { $count += ([Convert]::ToString([int]$p,2).ToCharArray() | Where-Object{$_ -eq '1'}).Count }
$prefixLength = $count

try {
    Get-NetIPAddress -InterfaceAlias $interfaceAlias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
    Get-NetRoute -InterfaceAlias $interfaceAlias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue
    New-NetIPAddress -InterfaceAlias $interfaceAlias -IPAddress $ipAddress -PrefixLength $prefixLength -DefaultGateway $gateway -ErrorAction Stop | Out-Null
    Write-Host "     [OK] IP地址设置成功" -ForegroundColor Green
    if ($dns1) { Set-DnsClientServerAddress -InterfaceAlias $interfaceAlias -ServerAddresses $dns1,$dns2 -ErrorAction Stop; Write-Host "     [OK] DNS设置成功" -ForegroundColor Green }
} catch { Write-Host "[错误] $($_.Exception.Message)" -ForegroundColor Red; exit 1 }

Write-Host "[4] 设置计算机名称..." -ForegroundColor Yellow
if ($env:COMPUTERNAME -ne $targetConfig.Name) {
    try { Rename-Computer -NewName $targetConfig.Name -Force -ErrorAction Stop; Write-Host "     [OK] 计算机名已设置，重启生效" -ForegroundColor Green } 
    catch { Write-Host "[错误] $($_.Exception.Message)" -ForegroundColor Red }
} else { Write-Host "     [OK] 计算机名已是: $($targetConfig.Name)" -ForegroundColor Green }

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "配置完成!" -ForegroundColor Green