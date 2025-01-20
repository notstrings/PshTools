$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    Enum enmCompareMode {
        Name = 0
        MD5  = 1
    }
    class RemoveDupConf {
        [enmCompareMode] `$CompareMode
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object RemoveDupConf -Property @{
            CompareMode = [enmCompareMode]::Name
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [RemoveDupConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([RemoveDupConf]) $json
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

function local:RemoveDupFile([string[]] $Targets) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 本体処理
        $hash = @{}
        switch ($Conf.CompareMode) {
            "Name" {
                $Targets | ForEach-Object {
                    Get-ChildItem -LiteralPath $_ -File -Recurse |
                    ForEach-Object {
                        $uniqkey = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                        if (-not $hash.ContainsKey($uniqkey)){
                            $hash[$uniqkey] = @()
                        }
                        $hash[$uniqkey] += $_
                    }
                }
            }
            "MD5" {
                $Targets | ForEach-Object {
                    Get-ChildItem -LiteralPath $_ -File -Recurse |
                    ForEach-Object {
                        $uniqkey = Get-FileHash -LiteralPath $_.FullName -Algorithm MD5
                        if (-not $hash.ContainsKey($uniqkey)){
                            $hash[$uniqkey] = @()
                        }
                        $hash[$uniqkey] += $_
                    }
                }
            }
        }
        $hash.Values | ForEach-Object {
            $_ |
            Sort-Object -Property Length -Descending |
            Select-Object -Skip 1 | ForEach-Object {
                MoveTrush -Path $_.FullName
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb")

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
	# 引数確認
    if ($args.Length -eq 0) {
        EditConfFile $Title $ConfPath
        exit
    }
	# 処理実行
    RemoveDupFile $args
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
