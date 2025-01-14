$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Enum enmCompareMode {
    Name = 0
    MD5  = 1
}
class Conf {
    [enmCompareMode] $CompareMode
}

# 設定初期化
function local:InitConf([string] $sPath) {
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $conf = New-Object Conf -Property @{
            CompareMode = [enmCompareMode]::Name
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
    $conf = LoadConf $ConfPath
    $ret = ShowSettingDialog $Title $conf
    if ($ret -eq "OK") {
        SaveConf $ConfPath $conf
    }
}

## 本体 #######################################################################

function local:RemoveDupFile([string[]] $Targets) {
    $hash = @{}
    switch ($conf.CompareMode) {
        "Name" {
    $Targets | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -File |
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
                Get-ChildItem -LiteralPath $_ -File |
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
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb")

try {
    $null = Write-Host "---$Title---"
    # 設定取得
    InitConf $ConfPath
    $conf = LoadConf $ConfPath
	# 引数確認
    if ($args.Length -eq 0) {
        EditConf $ConfPath
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
