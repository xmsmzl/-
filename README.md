# -
因为懒，所以写一个小工具
<#
.SYNOPSIS
    System Ecosystem Collector V4 (Ultimate IE Deep Search，这名字专业不？)
    - Fixes missing IE items
    - Fixes "Unknown Control" by cross-referencing WOW6432Node and DLL metadata
	  - 请给作者一键三连~~~
#>

$ErrorActionPreference = "SilentlyContinue"

# 获取脚本当前路径
$ScriptPath = $PSScriptRoot
if (-not $ScriptPath) { $ScriptPath = $PWD.Path }
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$FilePath = "$ScriptPath\本机生态信息_$Timestamp.xlsx"

Write-Host "正在启动生态收集小工具（power by Zorro_Mao），初始化 Excel 中..." -ForegroundColor Cyan

# 创建 Excel
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Add()
}
catch {
    Write-Host "错误: 无法启动 Excel。" -ForegroundColor Red
    Pause; Exit
}

# 清理 Sheet
while ($Workbook.Sheets.Count -gt 1) { $Workbook.Sheets.Item(2).Delete() }

function Write-ToSheet {
    param ([string]$SheetName, [array]$Data, [array]$Headers)
    if ($Workbook.Sheets.Item(1).Name -match "Sheet") { $Worksheet = $Workbook.Sheets.Item(1) } 
    else { $Worksheet = $Workbook.Sheets.Add() }
    $Worksheet.Name = $SheetName
    
    # 格式化表头
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Worksheet.Cells.Item(1, $i + 1) = $Headers[$i]
        $Worksheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $Worksheet.Cells.Item(1, $i + 1).Interior.ColorIndex = 37
        $Worksheet.Cells.Item(1, $i + 1).Borders.LineStyle = 1
    }

    $Row = 2
    foreach ($Item in $Data) {
        $Col = 1
        foreach ($Header in $Headers) {
            $Worksheet.Cells.Item($Row, $Col) = $Item.$Header
            $Col++
        }
        $Row++
    }
    $Worksheet.Columns.AutoFit() | Out-Null
}

# ========================================================
# 1. 软件生态 (标准扫描)
# ========================================================
Write-Host "1/3 正在扫描已安装软件..."
$SoftwareList = @()
$UninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
)
foreach ($Key in $UninstallKeys) {
    Get-ChildItem $Key -ErrorAction SilentlyContinue | ForEach-Object {
        $Props = Get-ItemProperty $_.PSPath
        if ($Props.DisplayName) {
            $SoftwareList += [PSCustomObject]@{
                '软件名称' = $Props.DisplayName
                '版本'     = $Props.DisplayVersion
                '开发商'   = $Props.Publisher
                '安装路径' = $Props.InstallLocation
            }
        }
    }
}
$SoftwareList = $SoftwareList | Sort-Object '软件名称' -Unique
Write-ToSheet -SheetName "1-软件生态" -Data $SoftwareList -Headers @("软件名称", "版本", "开发商", "安装路径")

# ========================================================
# 2. 外设 (WMI 扫描)
# ========================================================
Write-Host "2/3 正在扫描外设硬件..."
$Peripherals = Get-CimInstance Win32_PnPEntity | Where-Object { 
    $_.Status -eq "OK" -and 
    ($_.PNPClass -match "USB|Image|Printer|Mouse|Keyboard|HID|Biometric|SmartCard") 
} 
$PeripheralsExport = $Peripherals | ForEach-Object {
    [PSCustomObject]@{
        '设备名称' = $_.Name
        '类别'     = $_.PNPClass
        '厂商'     = $_.Manufacturer
        '设备ID'   = $_.DeviceID
    }
}
Write-ToSheet -SheetName "2-外设列表" -Data $PeripheralsExport -Headers @("设备名称", "类别", "厂商", "设备ID")

# ========================================================
# 3. 浏览器插件 (IE 强力修复版，没办法只能这样绕路)
# ========================================================
Write-Host "3/3 正在读取浏览器插件..."
$AllExtensions = @()

# --- A. Chromium (Edge/Chrome) 本地化名称修复逻辑 ---
$Browsers = @{ "Edge"="$env:LOCALAPPDATA\Microsoft\Edge\User Data"; "Chrome"="$env:LOCALAPPDATA\Google\Chrome\User Data" }
foreach ($BrowserName in $Browsers.Keys) {
    $UserDataPath = $Browsers[$BrowserName]
    if (Test-Path $UserDataPath) {
        $Profiles = Get-ChildItem $UserDataPath -Directory | Where-Object { $_.Name -match "^(Default|Profile)" }
        foreach ($Profile in $Profiles) {
            $ExtRoot = Join-Path $Profile.FullName "Extensions"
            if (Test-Path $ExtRoot) {
                $Extensions = Get-ChildItem $ExtRoot -Directory
                foreach ($Ext in $Extensions) {
                    $VerDir = Get-ChildItem $Ext.FullName -Directory | Sort-Object Name -Descending | Select-Object -First 1
                    $Manifest = Join-Path $VerDir.FullName "manifest.json"
                    if (Test-Path $Manifest) {
                        try {
                            $Json = Get-Content $Manifest -Raw -Encoding UTF8 | ConvertFrom-Json
                            $Name = $Json.name
                            $Ver  = $Json.version
                            # 解析 __MSG_
                            if ($Name -match "__MSG_(.+?)__") {
                                $MsgKey = $Matches[1]
                                $Locale = if ($Json.default_locale) { $Json.default_locale } else { "en" }
                                $MsgPath = Join-Path $VerDir.FullName "_locales\zh_CN\messages.json"
                                if (-not (Test-Path $MsgPath)) { $MsgPath = Join-Path $VerDir.FullName "_locales\$Locale\messages.json" }
                                if (Test-Path $MsgPath) {
                                    try { $MsgJson = Get-Content $MsgPath -Raw -Encoding UTF8 | ConvertFrom-Json; if ($MsgJson.$MsgKey.message) { $Name = $MsgJson.$MsgKey.message } } catch {}
                                }
                            }
                            $AllExtensions += [PSCustomObject]@{ '浏览器'="$BrowserName"; '类型'="Extension"; '名称'=$Name; '版本'=$Ver; '标识(CLSID/ID)'=$Ext.Name; '文件/DLL路径'="Profile: $($Profile.Name)" }
                        } catch {}
                    }
                }
            }
        }
    }
}

# --- B. IE 插件 (V4 核心：双向查找 + DLL 强制解析)（没错，已经是第四版了） ---

# 定义函数：通过 CLSID 查找真实名称
function Get-ClsidRealName {
    param ($Clsid)
    $ResolvedName = $null
    $ResolvedPath = $null

    # 构造可能的注册表位置 (64位 和 32位 都要找)
    $SearchPaths = @(
        "HKCR:\CLSID\$Clsid",
        "HKCR:\WOW6432Node\CLSID\$Clsid",
        "HKLM:\SOFTWARE\Classes\CLSID\$Clsid",
        "HKLM:\SOFTWARE\Classes\WOW6432Node\CLSID\$Clsid",
        "HKCU:\Software\Classes\CLSID\$Clsid"
    )

    foreach ($Path in $SearchPaths) {
        if (Test-Path $Path) {
            # 1. 尝试直接获取注册表名称
            $RegVal = (Get-ItemProperty $Path -ErrorAction SilentlyContinue)."(default)"
            if ($RegVal -and $RegVal -ne "" -and $RegVal -notmatch "^{") { 
                if (-not $ResolvedName) { $ResolvedName = $RegVal }
            }

            # 2. 寻找 InprocServer32 (DLL 路径)
            $DllKey = Join-Path $Path "InprocServer32"
            if (Test-Path $DllKey) {
                $DllVal = (Get-ItemProperty $DllKey -ErrorAction SilentlyContinue)."(default)"
                if ($DllVal) {
                    $DllVal = $DllVal.Replace('"', '') # 去除引号
                    $DllVal = [Environment]::ExpandEnvironmentVariables($DllVal) # 转换 %SystemRoot%
                    if (Test-Path $DllVal) {
                        $ResolvedPath = $DllVal
                        # 3. 终极技：读取 DLL 文件信息
                        try {
                            $VerInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($DllVal)
                            if ($VerInfo.FileDescription -and $VerInfo.FileDescription -ne "") {
                                $ResolvedName = $VerInfo.FileDescription # 文件描述通常是最准确的名称 (如 Adobe PDF Helper)
                                break # 找到最准确的就停止
                            }
                            if (-not $ResolvedName -and $VerInfo.ProductName) {
                                $ResolvedName = $VerInfo.ProductName
                            }
                        } catch {}
                    }
                }
            }
        }
    }
    return @{ Name = $ResolvedName; Path = $ResolvedPath }
}

# 遍历 IE BHO 和 Toolbar 的位置
$IE_Scan_Locations = @(
    @{ Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"; Type="IE BHO (Machine)" },
    @{ Path="HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"; Type="IE BHO (32-bit)" },
    @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"; Type="IE BHO (User)" },
    @{ Path="HKLM:\SOFTWARE\Microsoft\Internet Explorer\Toolbar"; Type="IE Toolbar (Machine)" },
    @{ Path="HKLM:\SOFTWARE\WOW6432Node\Microsoft\Internet Explorer\Toolbar"; Type="IE Toolbar (32-bit)" },
    @{ Path="HKCU:\Software\Microsoft\Internet Explorer\Toolbar"; Type="IE Toolbar (User)" }
)

foreach ($Loc in $IE_Scan_Locations) {
    if (Test-Path $Loc.Path) {
        # 获取下面的所有子项 (即 CLSID)
        $Items = Get-ChildItem $Loc.Path -ErrorAction SilentlyContinue
        foreach ($Item in $Items) {
            $Clsid = $Item.PSChildName
            
            # 调用解析函数
            $Info = Get-ClsidRealName -Clsid $Clsid
            
            $FinalName = if ($Info.Name) { $Info.Name } else { "未知控件/无描述" }
            $FinalPath = if ($Info.Path) { $Info.Path } else { "未找到 DLL 文件" }

            $AllExtensions += [PSCustomObject]@{
                '浏览器'       = "Internet Explorer"
                '类型'         = $Loc.Type
                '名称'         = $FinalName
                '版本'         = "N/A"
                '标识(CLSID/ID)' = $Clsid
                '文件/DLL路径'   = $FinalPath
            }
        }
    }
}

# 仅根据 CLSID 去重，保留最详细的一个
$AllExtensions = $AllExtensions | Sort-Object '标识(CLSID/ID)' -Unique
Write-ToSheet -SheetName "3-浏览器插件" -Data $AllExtensions -Headers @("浏览器", "类型", "名称", "版本", "标识(CLSID/ID)", "文件/DLL路径")

# ========================================================
# 保存退出
# ========================================================
$Workbook.SaveAs($FilePath)
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()

Write-Host "`n----------------------------------------" -ForegroundColor Green
Write-Host "生态采集完成！"
Write-Host "清单保存路径: $FilePath" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Green
Pause
