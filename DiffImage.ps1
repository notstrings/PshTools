﻿$ErrorActionPreference = "Stop"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Enum enmGravityType {
    NorthWest = 0
    North     = 1
    NorthEast = 2
    West      = 3
    Center    = 4
    East      = 5
    SouthWest = 6
    South     = 7
    SouthEast = 8
}
class Conf {
    [bool]            $FitSize
    [enmGravityType]  $Align
}

# 設定初期化
function local:InitConf([string] $sPath) {
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $conf = New-Object Conf -Property @{
            FitSize = $true
            Align   = [enmGravityType]::Center
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
    $conf = GenClassByPSCustomObject ([Conf]) $json
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

function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install imagemagick
}

function local:DiffImage([System.IO.FileInfo] $LHS, [System.IO.FileInfo] $RHS) {
    try {
        $IMPath = "magick.exe"
        $LSrcPath = $LHS.FullName
        $RSrcPath = $RHS.FullName
        $opt1 = ""
        if ($conf.FitSize -eq $true) {
            $opt1 += "-resize 800x800 "
        }
        $opt2 = $conf.Align.ToString()
        if ($LHS.LastWriteTime -le $RHS.LastWriteTime) {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$LSrcPath"" $opt1 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$RSrcPath"" $opt1 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        } else {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$RSrcPath"" $opt1 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$LSrcPath"" $opt1 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        }
        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""tempLHS.png"" ""tempRHS.png"" -compose Multiply -gravity $opt2 -composite ""diff.png"""
        $null = Start-Process -NoNewWindow -Wait -FilePath "mspaint.exe" -ArgumentList "diff.png"
    } finally {
        $null = Remove-Item -Path @("diff.png", "tempLHS.png", "tempRHS.png") -Force
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.png", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb.png")

try {
    $null = Write-Host "---$Title---"
    # 設定取得
    InitConf $ConfPath
    $conf = LoadConf $ConfPath
	# 引数確認
    if ($args.Count -ne 2) {
        EditConf $ConfPath
        exit
    }
	# 処理実行
    $exist = $true
    $exist = $exist -and (Test-Path -LiteralPath $args[0])
    $exist = $exist -and (Test-Path -LiteralPath $args[1])
    if ($exist) {
        DiffImage (Get-Item $args[0]) (Get-Item $args[1])
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
