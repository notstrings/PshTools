$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.Design

Invoke-Expression -Command @"
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
    class DiffImageConf {
        [bool]            `$FitSize
        [enmGravityType]  `$Align
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object DiffImageConf -Property @{
            FitSize = $true
            Align   = [enmGravityType]::Center
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [DiffImageConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([DiffImageConf]) $json
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

function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install imagemagick
}

function local:DiffImage([System.IO.FileInfo] $LHS, [System.IO.FileInfo] $RHS) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 本体処理
        $IMPath = "magick.exe"
        $LSrcPath = $LHS.FullName
        $RSrcPath = $RHS.FullName
        $opt1 = ""
        if ($Conf.FitSize -eq $true) {
            $opt1 += "-resize 800x800 "
        }
        $opt2 = $Conf.Align.ToString()
        if ($LHS.LastWriteTime -le $RHS.LastWriteTime) {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$LSrcPath"" $opt1 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$RSrcPath"" $opt1 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        } else {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$RSrcPath"" $opt1 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$LSrcPath"" $opt1 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        }
        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""tempLHS.png"" ""tempRHS.png"" -compose Multiply -gravity $opt2 -composite ""diff.png"""
        $null = Start-Process -NoNewWindow -Wait -FilePath "mspaint.exe" -ArgumentList "diff.png"
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    } finally {
        $null = Remove-Item -Path @("diff.png", "tempLHS.png", "tempRHS.png") -Force
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.png", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb.png")

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
	# 引数確認
    if ($args.Count -ne 2) {
        EditConfFile $Title $ConfPath
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
