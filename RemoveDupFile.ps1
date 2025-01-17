$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.Design

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
function local:InitConf([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $conf = New-Object RemoveDupConf -Property @{
            CompareMode = [enmCompareMode]::Name
        }
        SaveConf $Path $conf
    }
}
# 設定書込
function local:SaveConf([string] $Path, [RemoveDupConf] $conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConf([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $conf = ConvertFromPSCO ([RemoveDupConf]) $json
    return $conf
}
# 設定編集
function local:EditConf([string] $Title, [string] $Path) {
    $conf = LoadConf $Path
    $ret = ShowSettingDialog $Title $conf
    if ($ret -eq "OK") {
        SaveConf $Path $conf
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
        EditConf $Title $ConfPath
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
