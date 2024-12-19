$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function CheckFolderUpdate([string] $MonitorName, [string] $MonitorPath) {
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
    if ((Test-Path $PrevPath) -eq $false){
        $Ret = "初回動作"
    }else{
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
            if($_.SideIndicator -eq "<=") {
                if($null -ne $_.FullName){$DelFile += $_.FullName } 
            } elseif ($_.SideIndicator -eq "=>") {
                if($null -ne $_.FullName){$AddFile += $_.FullName} 
            } elseif ($_.SideIndicator -eq "==") {
                if($null -ne $_.FullName){$ModFile += $_.FullName} 
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

# 監視処理
function FolderMonitor() {
    $jsonString = Get-Content -LiteralPath "$($PSScriptRoot)\Config\MonitorSetting.json" -Raw
    $jsonObject = ConvertFrom-Json $jsonString
    $jsonObject.Monitors | ForEach-Object {
        $MonitorName = ($_.MonitorName)
        $MonitorPath = ($_.MonitorPath -replace "\\", "\")
        # フォルダ更新検出
        $Result = CheckFolderUpdate $MonitorName $MonitorPath
        # 自分にIPMessengerを投げる
        if ($Result -ne ""){
            SendIPMsg -Message "$($MonitorName)監視\n$($MonitorPath)\n$($Result)" 
        }
    }
}

class ConfChild {
    [System.ComponentModel.Description("監視名称")]
    [string]$MonName
    [System.ComponentModel.Description("監視位置")]
    [string]$MonPath
}
class Conf {
    [ConfChild[]]$ConfChild
}

function FunctionName() {
    $conf = New-Object Conf
    $conf.ConfChild = @()
    $ret = ShowSettingDialog "Title" $conf
    if ($ret[0] -eq "OK") {
        $ret[1]
    }    
}

FunctionName


# 常駐監視
# RunInTray "Monitor" 0x0000ff { FunctionName } { FolderMonitor } (5*60*1000)
