$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

## 設定関連 #######################################################################

class ConfChild {
    [System.ComponentModel.Description("監視名称")]
    [string]$MonName
    [System.ComponentModel.Description("監視位置")]
    [string]$MonPath
}
class Conf {
    [ConfChild[]]$ConfChild
}
$conf = New-Object Conf
$conf.ConfChild = @()

# 設定書込
function local:SaveConf([string] $sPath, [Conf] $conf) {
    $conf.ConfChild | ForEach-Object {
        $_.MonPath = $_.MonPath -replace "\\", "/"
    }
    $conf | ConvertTo-Json | Out-File -FilePath $sPath
}
# 設定読出
function local:LoadConf([string] $sPath) {
    # デフォ値生成※設定ファイル無しの場合
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $child1 = New-Object ConfChild -Property @{MonName = "監視01"; MonPath = ""}
        $child2 = New-Object ConfChild -Property @{MonName = "監視02"; MonPath = ""}
        $child3 = New-Object ConfChild -Property @{MonName = "監視03"; MonPath = ""}
        $inst = New-Object Conf -Property @{ ConfChild = @($child1, $child2, $child3) }
        SaveConf $sPath $inst
    }
    # 設定読出
    $jcon = Get-Content -Path $sPath | ConvertFrom-Json
    $inst = New-Object Conf 
    $inst.ConfChild = $jcon.ConfChild | ForEach-Object {
        New-Object ConfChild -Property @{MonName = $_.MonName; MonPath = ($_.MonPath -replace "/", "\")}
    }
    return $inst
}
# 設定編集
function local:EditConf() {
    # 設定読出
    $crnt = LoadConf "$($PSScriptRoot)\Config\MonitorSetting.json"
    # 設定画面表示
    $ret = ShowSettingDialog "Title" $crnt
    if ($ret[0] -eq "OK") {
        # 設定書込
        SaveConf "$($PSScriptRoot)\Config\MonitorSetting.json" $crnt
    }
}
# EditConf

## 監視処理 #######################################################################

function local:FolderMonitor() {
    $Message = ""
    # 更新検出
    $conf = LoadConf "$($PSScriptRoot)\Config\MonitorSetting.json"
    $conf.ConfChild | ForEach-Object {
        $MonitorName = $_.MonName
        $MonitorPath = $_.MonPath
        if (Test-Path -LiteralPath $MonitorPath) {
            $Result = CheckFolderUpdate $MonitorName $MonitorPath
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
    Export-Csv -Path $CrntPath -NoTypeInformation -Encoding ([System.Text.Encoding]::GetEncoding("Shift_JIS"))
    
    # 昨今の監視フォルダ状況差分を確認
    if (Test-Path $PrevPath){
        # ファイル名＋更新日時の差分が監視フォルダ変更の全体像で...
        # ・追加は前回リストに無い差分
        # ・削除は今回リストに無い差分
        # ・更新は更新日時のみの差分
        $Diff = DiffContent -LHSPath $PrevPath -RHSPath $CrntPath -Encoding ([System.Text.Encoding]::GetEncoding("Shift_JIS"))
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
        $DelFile | ForEach-Object { $Ret += "DEL:$([System.IO.Path]::GetRelativePath($MonitorPath, $_))\n" } | Out-Null
        $AddFile | ForEach-Object { $Ret += "ADD:$([System.IO.Path]::GetRelativePath($MonitorPath, $_))\n" } | Out-Null
        $ModFile | ForEach-Object { $Ret += "MOD:$([System.IO.Path]::GetRelativePath($MonitorPath, $_))\n" } | Out-Null
    }

    # 現在のフォルダ状況を過去のフォルダ状況とする
    $null = Copy-Item -Destination $PrevPath -LiteralPath $CrntPath -Force

    return $Ret
}

# 常駐監視
RunInTray "Monitor" 0x0000ff { EditConf } { FolderMonitor } (5*60*1000)
