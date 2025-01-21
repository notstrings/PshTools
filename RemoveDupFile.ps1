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
    Enum enmReduceMode {
        New   = 0
        Old   = 1
        Small = 2
        Large = 3
    }
    class RemoveDupConf {
        [enmCompareMode] `$CompareMode
        [enmReduceMode]  `$ReduceMode
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object RemoveDupConf -Property @{
            CompareMode = [enmCompareMode]::Name
            ReduceMode  = [enmReduceMode]::New
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

function local:ReduceDupFile([string[]] $Targets) {
    # 設定取得
    $Conf = LoadConfFile $ConfPath
    # 本体処理
    $Hash = @{}
    foreach ($Target in $Targets) {
        if (Test-Path -LiteralPath $Target) {
            if ([System.IO.Directory]::Exists($Target)) {
                MakeHashDir  $Hash (Get-Item $Target) $Conf.CompareMode
            } else {
                MakeHashFile $Hash (Get-Item $Target) $Conf.CompareMode
            }
        }
    }
    RemoveDupFile $Hash $Conf.ReduceMode
}
function local:MakeHashDir([hashtable] $Hash, [System.IO.DirectoryInfo] $Target, [enmCompareMode] $CompareMode) {
    Get-ChildItem -LiteralPath $Target.FullName -File -Recurse | ForEach-Object {
        MakeHashFile $Hash $_ $CompareMode
    }
}
function local:MakeHashFile([hashtable] $Hash, [System.IO.FileInfo] $Target, [enmCompareMode] $CompareMode) {
    switch ($CompareMode) {
        "Name" {
            $uniqkey = [System.IO.Path]::GetFileNameWithoutExtension($Target.FullName)
            $uniqkey = RestrictTextZen   -Text $uniqkey -Chars "Ａ-Ｚａ-ｚ０-９　（）［］｛｝"
            $uniqkey = RestrictTextHan   -Text $uniqkey
            $uniqkey = RestrictTextBlank -Text $uniqkey
            $uniqkey = RemoveAllBrackets -Text $uniqkey
            $uniqkey = $uniqkey.ToLower()
        }
        "MD5" {
            $uniqkey = Get-FileHash -LiteralPath $Target.FullName -Algorithm MD5
        }
    }
    if (-not $Hash.ContainsKey($uniqkey)){
        $Hash[$uniqkey] = @()
    }
    $Hash[$uniqkey] += $Target
}
function local:RemoveDupFile([hashtable] $Hash, [enmReduceMode] $ReduceMode) {
    switch ($ReduceMode) {
        "New" {
            $Hash.Values | ForEach-Object {
                $_ | 
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -Skip 1 | ForEach-Object {
                    MoveTrush -Path $_.FullName
                }
            }
        }
        "Old" {
            $Hash.Values | ForEach-Object {
                $_ | 
                Sort-Object -Property LastWriteTime |
                Select-Object -Skip 1 | ForEach-Object {
                    MoveTrush -Path $_.FullName
                }
            }
        }
        "Small" {
            $Hash.Values | ForEach-Object {
                $_ | 
                Sort-Object -Property Length |
                Select-Object -Skip 1 | ForEach-Object {
                    MoveTrush -Path $_.FullName
                }
            }
        }
        "Large" {
            $Hash.Values | ForEach-Object {
                $_ | 
                Sort-Object -Property Length -Descending |
                Select-Object -Skip 1 | ForEach-Object {
                    MoveTrush -Path $_.FullName
                }
            }
        }
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
    ReduceDupFile $args
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
