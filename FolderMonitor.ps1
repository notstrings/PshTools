$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
	winget install "FastCopy.IPMsg"
}

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    class FolderMonitorConf {
        [MonitorTarget[]] `$MonitorTargets
        [System.ComponentModel.Description("変更の適用に再起動が必要")]
        [int] `$Interval
    }
    class MonitorTarget {
        [System.ComponentModel.Description("監視名称")]
        [string] `$MonName
        [System.ComponentModel.Description("監視位置")]
        [string] `$MonPath
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $child1 = New-Object MonitorTarget -Property @{MonName = "Mon01"; MonPath = ""}
        $child2 = New-Object MonitorTarget -Property @{MonName = "Mon02"; MonPath = ""}
        $child3 = New-Object MonitorTarget -Property @{MonName = "Mon03"; MonPath = ""}
        $Conf = New-Object FolderMonitorConf -Property @{
            MonitorTargets = @($child1, $child2, $child3)
            Interval = (5*60*1000)
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [FolderMonitorConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([FolderMonitorConf]) $json
    return $Conf
}
# 設定編集
function local:EditConfFile([string] $Title, [string] $Path) {
    $Conf = LoadConfFile $Path
    $ret = ShowSettingDialog $Title $Conf
    if ($ret -eq "OK") {
        SaveConfFile $Path $Conf
    }
}

## 本体 #######################################################################

function local:FolderMonitor() {
    # 設定取得
    $Conf = LoadConfFile $ConfPath
    # 更新検出
    $Result = ""
    $Conf.MonitorTargets | ForEach-Object {
        $MonitorName = $_.MonName
        $MonitorPath = $_.MonPath
        if ( ("" -ne $MonitorPath) -and (Test-Path -LiteralPath $MonitorPath)) {
            $Result += CheckFolderUpdate $MonitorName $MonitorPath $MonitorInterval
        }
    }
    # 結果表示
    if ("" -ne $Result ){
        SendIPMsg -Message $Result
    }
}
function local:CheckFolderUpdate([string] $MonitorName, [string] $MonitorPath) {
    $Ret = ""

    # ワーキングフォルダを確保
    $PrevPath = "$($PSScriptRoot)\Monitor\$($MonitorName)Prev.csv"
    $CrntPath = "$($PSScriptRoot)\Monitor\$($MonitorName)Crnt.csv"
    $null = New-Item ([System.IO.Path]::GetDirectoryName($PrevPath)) -ItemType Directory -ErrorAction SilentlyContinue
    $null = New-Item ([System.IO.Path]::GetDirectoryName($CrntPath)) -ItemType Directory -ErrorAction SilentlyContinue

    # 現在の監視フォルダ状況を取得
    Get-ChildItem $MonitorPath -File -Recurse |
    Select-Object FullName, LastWriteTime |
    Export-Csv -Path $CrntPath -NoTypeInformation -Encoding "OEM" # PowerShell5互換のためエンコーディングを文字指定

    # 昨今の監視フォルダ状況差分を確認
    if (Test-Path $PrevPath){
        # ファイル名と変更日時の差分を抽出
        $PrevAllCSV = @(Import-Csv -LiteralPath $PrevPath -Encoding "OEM") # PowerShell5互換のためエンコーディングを文字指定
        $CrntAllCSV = @(Import-Csv -LiteralPath $CrntPath -Encoding "OEM") # PowerShell5互換のためエンコーディングを文字指定
        $PrevDifCSV = @()
        $CrntDifCSV = @()
        Compare-Object -ReferenceObject $PrevAllCSV -DifferenceObject $CrntAllCSV -Property FullName, LastWriteTime |
        ForEach-Object {
            if($_.SideIndicator -eq "<=") {
                $PrevDifCSV += [PSCustomObject]@{FullName = $_.FullName; LastWriteTime = $_.LastWriteTime}
            } elseif ($_.SideIndicator -eq "=>") {
                $CrntDifCSV += [PSCustomObject]@{FullName = $_.FullName; LastWriteTime = $_.LastWriteTime}
            }
        } | Out-Null
        # ファイル名差分だけで絞り込んで追加/削除/変更を識別
        $ModFile = @()
        $DelFile = @()
        $AddFile = @()
        Compare-Object -ReferenceObject $PrevDifCSV -DifferenceObject $CrntDifCSV -IncludeEqual -Property FullName |
        ForEach-Object {
            if($null -ne $_.FullName -and "FullName" -ne $_.FullName){
                if($_.SideIndicator -eq "<=") {
                    $DelFile += $_.FullName
                } elseif ($_.SideIndicator -eq "=>") {
                    $AddFile += $_.FullName
                } elseif ($_.SideIndicator -eq "==") {
                    $ModFile += $_.FullName
                }
            }
        } | Out-Null
        # 結果出力
        $AddFile | ForEach-Object { $Ret += "ADD:$($_.Replace($MonitorPath,'.'))\n" } | Out-Null
        $DelFile | ForEach-Object { $Ret += "DEL:$($_.Replace($MonitorPath,'.'))\n" } | Out-Null
        $ModFile | ForEach-Object { $Ret += "MOD:$($_.Replace($MonitorPath,'.'))\n" } | Out-Null
        if ($Ret -ne ""){
            $Ret = "---\nMonitorName:$MonitorName\nMonitorPath:$MonitorPath\n" + $Ret
        }
    }

    # 現在のフォルダ状況を過去のフォルダ状況とする
    $null = Copy-Item -Destination $PrevPath -LiteralPath $CrntPath -Force

    return $Ret
}
# FolderMonitor

###############################################################################

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
	# 処理実行
    $Conf = LoadConfFile $ConfPath
    RunInTaskTray $Title 0x0000ff { EditConfFile $Title $ConfPath } { FolderMonitor } $Conf.Interval
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
