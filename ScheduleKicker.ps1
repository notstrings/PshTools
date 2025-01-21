$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    class ScheduleKickerConf {
        [ScheduleTask[]] `$ScheduleTasks
    }
    class ScheduleTask {
        [string] `$ScheduleName
        [string] `$ScheduleTime
        [string] `$SchedulePath
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $child1 = New-Object ScheduleTask -Property @{ScheduleName = "Task01"; ScheduleTime = "12:00"; SchedulePath = ""}
        $child2 = New-Object ScheduleTask -Property @{ScheduleName = "Task02"; ScheduleTime = "15:00"; SchedulePath = ""}
        $child3 = New-Object ScheduleTask -Property @{ScheduleName = "Task03"; ScheduleTime = "17:00"; SchedulePath = ""}
        $Conf = New-Object ScheduleKickerConf -Property @{
            ScheduleTasks = @($child1, $child2, $child3)
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [ScheduleKickerConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([ScheduleKickerConf]) $json
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

function local:ScheduleKicker() {
    # 設定取得
    $Conf = LoadConfFile $ConfPath
    # 更新検出
    $global:CrntRunTime = Get-Date
    $Conf.ScheduleTasks | ForEach-Object {
        $ScheduleName = $_.ScheduleName
        $ScheduleTime = $_.ScheduleTime
        $SchedulePath = $_.SchedulePath
        if (CheckTime $global:PrevRunTime $global:CrntRunTime $ScheduleTime) {
            ShowToast $Title $ScheduleName $SchedulePath
            if ("" -eq $SchedulePath) {
                Invoke-Expression -Command $SchedulePath
            }
        }
    }
    $global:PrevRunTime = $global:CrntRunTime
}
function local:CheckTime([datetime] $PrevRunTime, [datetime] $CrntRunTime, [string] $CheckTime) {
    $Ret = $false
    $TargetTime = [datetime]::ParseExact($CheckTime, "HH:mm", $null)
    if ($PrevRunTime.TimeOfDay -le $TargetTime.TimeOfDay -and $CrntRunTime.TimeOfDay -ge $TargetTime.TimeOfDay) {
        $Ret = $true
    }
    return $Ret
}

###############################################################################

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
	# 処理実行
    $global:CrntRunTime = Get-Date
    $global:PrevRunTime = Get-Date
    RunInTaskTray $Title 0x00ff00 { EditConfFile $Title $ConfPath } { ScheduleKicker } (1000)
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
