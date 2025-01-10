$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
	winget install "FastCopy.IPMsg"
}

## 設定 #######################################################################

class Conf {
    [ConfChild[]] $ConfChild
    [System.ComponentModel.Description("変更の適用に再起動が必要")]
    [int]         $Interval
}
class ConfChild {
    [System.ComponentModel.Description("監視名称")]
    [string]$MonName
    [System.ComponentModel.Description("監視位置")]
    [string]$MonPath
}

# 設定初期化
function local:InitConf([string] $sPath) {
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $child1 = New-Object ConfChild -Property @{MonName = "Mon01"; MonPath = ""}
        $child2 = New-Object ConfChild -Property @{MonName = "Mon02"; MonPath = ""}
        $child3 = New-Object ConfChild -Property @{MonName = "Mon03"; MonPath = ""}
        $conf = New-Object Conf -Property @{
            ConfChild = @($child1, $child2, $child3)
            Interval = (5*60*1000)
        }
        SaveConf $sPath $conf
    }
}
# 設定書込
function local:SaveConf([string] $sPath, [Conf] $conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($sPath)) -ItemType Directory -ErrorAction SilentlyContinue
    $conf | ConvertTo-Json | Out-File -FilePath $sPath
}
# 設定読出
function local:LoadConf([string] $sPath) {
    $json = Get-Content -Path $sPath | ConvertFrom-Json
    $conf = ConvertFromPSCO ([Conf]) $json
    return $conf
}
# 設定編集
function local:EditConf([string] $sPath) {
    $conf = LoadConf $sPath
    $ret = ShowSettingDialog $Title $conf
    if ($ret -eq "OK") {
        SaveConf $sPath $conf
    }
}

## 本体 #######################################################################

function local:FolderMonitor() {
    $Message = ""
    # 更新検出
    $conf = LoadConf $ConfPath
    $conf.ConfChild | ForEach-Object {
        $MonitorName = $_.MonName
        $MonitorPath = $_.MonPath
        if ( ("" -ne $MonitorPath) -and (Test-Path -LiteralPath $MonitorPath)) {
            $Result = CheckFolderUpdate $MonitorName $MonitorPath $MonitorInterval
            if ($Result){
                $Message += "$($MonitorName)\n$($MonitorPath)\n$($Result)" 
            }
        }
    }
    # 結果表示
    if ("" -ne $Message ){
        SendIPMsg -Message $Message 
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
        # ファイル名＋更新日時の差分が監視フォルダ変更の全体像で...
        # ・追加は前回リストに無い差分
        # ・削除は今回リストに無い差分
        # ・更新は更新日時のみの差分
        $Diff = DiffContent -LHSPath $PrevPath -RHSPath $CrntPath -Encoding "OEM"  # PowerShell5互換のためエンコーディングを文字指定
        $LHSOnlyCSV = ($Diff[1] -join "`r`n") | ConvertFrom-Csv -Header "FullName", "LastModifyDate"
        $RHSOnlyCSV = ($Diff[2] -join "`r`n") | ConvertFrom-Csv -Header "FullName", "LastModifyDate"
        if ($LHSOnlyCSV.Length -eq 0){$LHSOnlyCSV = ""}
        if ($RHSOnlyCSV.Length -eq 0){$RHSOnlyCSV = ""}
        $ModFile = @()
        $DelFile = @()
        $AddFile = @()
        Compare-Object -ReferenceObject $LHSOnlyCSV -DifferenceObject $RHSOnlyCSV -IncludeEqual -Property FullName |
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
        $DelFile | ForEach-Object { $Ret += "DEL:$($_.Replace($MonitorPath,'.\'))\n" } | Out-Null
        $AddFile | ForEach-Object { $Ret += "ADD:$($_.Replace($MonitorPath,'.\'))\n" } | Out-Null
        $ModFile | ForEach-Object { $Ret += "MOD:$($_.Replace($MonitorPath,'.\'))\n" } | Out-Null
    }

    # 現在のフォルダ状況を過去のフォルダ状況とする
    $null = Copy-Item -Destination $PrevPath -LiteralPath $CrntPath -Force

    return $Ret
}
# FolderMonitor

###############################################################################

try {
    $null = Write-Host "---$Title---"
    # 設定取得
    InitConf $ConfPath
    $crnt = LoadConf $ConfPath
	# 処理実行
    RunInTaskTray $Title 0x0000ff { EditConf $ConfPath } { FolderMonitor } ($crnt.Interval)
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
