﻿$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    class ReduceDirConf {
        [string[]] `$RemFolders
        [string[]] `$RemFiles
        [string[]] `$RemExts
        [bool]     `$RemBlankFolder
        [bool]     `$RedOrphanFolder
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object ReduceDirConf -Property @{
            RemFolders      = @()
            RemFiles        = @("Thumbs.db",".DS_Store")
            RemExts         = @(".bak",".tmp")
            RemBlankFolder  = $true
            RedOrphanFolder = $false
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [ReduceDirConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([ReduceDirConf]) $json
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

# ファイル
function local:ReduceFile([System.IO.FileInfo] $Target) {
    Reduce $Target.FullName $false
}

# フォルダ
function local:ReduceDir([System.IO.DirectoryInfo] $Target) {
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ReduceFile $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ReduceDir  $elm
    }
    Reduce $Target.FullName $true
}

# 整理
function local:Reduce([string]$Target, [bool]$isDir) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 本体処理
        if ($isDir) {
            # フォルダ
            ## 不要フォルダ
            if ($null -ne $Conf.RemFolders){
                $remfolders = $Conf.RemFolders | ForEach-Object {$_.ToLower()}
                if ($remfolders -contains ([System.IO.Path]::GetFileName($Target).ToLower())) {
                    MoveTrush -Path $Target
                    return
                }
            }
            ## 空フォルダ
            if ($Conf.RemBlankFolder -eq $true){
                if (@(Get-ChildItem -LiteralPath $Target -File     ).Length -eq 0 -and
                    @(Get-ChildItem -LiteralPath $Target -Directory).Length -eq 0 ) {
                    MoveTrush -Path $Target
                    return
                }
            }
            ## 孤立フォルダ(引き上げ)
            if ($Conf.RedOrphanFolder -eq $true){
                if (@(Get-ChildItem -Path ([System.IO.Path]::Combine($Target, "..")) -File     ).Length -eq 0 -and
                    @(Get-ChildItem -Path ([System.IO.Path]::Combine($Target, "..")) -Directory).Length -eq 1 ) {
                    $dup = ([System.IO.Path]::Combine($Target, [System.IO.Path]::GetFileName($Target)))
                    if (Test-Path -LiteralPath $dup) {
                        $tmp = GenUniqName ($dup + "_") $true
                        Move-Item -LiteralPath $dup $tmp -Force
                        Move-Item -Path ($Target+"/*") ($Target+"/..") -Force
                        MoveTrush -Path $Target
                        Move-Item -LiteralPath $tmp $dup -Force
                    } else {
                        Move-Item -Path ($Target+"/*") ($Target+"/..") -Force
                        MoveTrush -Path $Target
                    }
                    return
                }
            }
        } else {
            # ファイル名
            ## 不要ファイル
            if ($null -ne $Conf.RemFiles){
                $remfiles = $Conf.RemFiles | ForEach-Object {$_.ToLower()}
                if ($remfiles -contains ([System.IO.Path]::GetFileName($Target).ToLower())) {
                    MoveTrush -Path $Target
                    return
                }
            }
            ## 不要拡張子
            if ($null -ne $Conf.RemExts){
                $remexts = $Conf.RemExts | ForEach-Object {$_.ToLower()}
                if ($remexts -contains ([System.IO.Path]::GetExtension($Target).ToLower())) {
                    MoveTrush -Path $Target
                    return
                }
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

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
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                ReduceDir (Get-Item $arg)
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
