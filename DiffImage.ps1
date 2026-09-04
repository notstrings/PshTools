$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

function local:Setup() {
    if (-not (Get-Command scoop -ErrorAction SilentlyContinue)) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install imagemagick
}

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
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
        [bool]           `$FitSize
        [enmGravityType] `$Align
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

function local:DiffImage([System.IO.FileInfo] $LHS, [System.IO.FileInfo] $RHS) {
    try {
        # フォルダ確認
        if ($LHS.DirectoryName -ne $RHS.DirectoryName) {
            throw "同じフォルダ内の画像を2枚選択してください"
        }
        # 本体処理
        $IMPath = "magick.exe"
        $SrcLHSPath = $LHS.FullName
        $SrcRHSPath = $RHS.FullName
        $DstBasePath = [System.IO.Path]::Combine("$($ENV:USERPROFILE)\Desktop", "Diff")
        $DstLHSPath = [System.IO.Path]::Combine($DstBasePath, "lhs_" + $LHS.BaseName + ".png")
        $DstRHSPath = [System.IO.Path]::Combine($DstBasePath, "rhs_" + $RHS.BaseName + ".png")
        $DstRSLPath = [System.IO.Path]::Combine($DstBasePath, "diff_" + $LHS.BaseName + "_" + $RHS.BaseName + ".png")
        $null = Write-Host "---"
        $null = Write-Host "lhs : $LHS"
        $null = Write-Host "rhs : $RHS"
        $null = Write-Host "rsl : $DstRSLPath"
        $null = New-Item $DstBasePath -ItemType Directory -ErrorAction SilentlyContinue
        $opt1 = ""
        if ($Conf.FitSize -eq $true) {
            $opt1 += "-resize 800x800 "
        }
        $opt2 = $Conf.Align.ToString()
        if ($LHS.LastWriteTime -le $RHS.LastWriteTime) {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$SrcLHSPath"" $opt1 -type GrayScale +level-colors Red,White  ""$DstLHSPath"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$SrcRHSPath"" $opt1 -type GrayScale +level-colors Blue,White ""$DstRHSPath"""
        } else {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$SrcRHSPath"" $opt1 -type GrayScale +level-colors Red,White  ""$DstLHSPath"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$SrcLHSPath"" $opt1 -type GrayScale +level-colors Blue,White ""$DstRHSPath"""
        }
        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$DstLHSPath"" ""$DstRHSPath"" -compose Multiply -gravity $opt2 -composite ""$DstRSLPath"""
        $null = Start-Process "$DstRSLPath"
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa\aaa.png", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa\bbb.png")

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
        $Conf = LoadConfFile $ConfPath
        DiffImage (Get-Item $args[0]) (Get-Item $args[1])
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
