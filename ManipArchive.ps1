$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install 7zip
}

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.Design

Invoke-Expression -Command @"
    Enum enmDivideType {
        None = 0
        New  = 1
        Old  = 2
    }
    class ManipArchiveConf {
        [bool]           `$Encrypt
        [enmDivideType]  `$DivideType
        [int]            `$DivideSize
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object ManipArchiveConf -Property @{
            Encrypt = $false
            DivideType = [enmDivideType]::None
            DivideSize = 1
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [ManipArchiveConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([ManipArchiveConf]) $json
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

# アーカイブファイル判定
function isArchive($Path) {
    return (([System.IO.Path]::GetExtension($Path).ToUpper()) -in @(".7Z", ".GZ", ".ZIP", ".BZ2", ".TAR", ".LZH", ".LZS", ".LHA", ".GZIP", ".LZMA"))
}

# 分割圧縮ファイル判定
function local:isDividedArchive($Path) {
    return (([System.IO.Path]::GetExtension($Path).ToUpper()) -match "\.\d{3}$")
}

# ランダムパスワード生成
function local:GetRandomPassword() {
    $text = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-+#$=_"
    $rand = New-Object System.Random
    return -join ((0..15) | ForEach-Object { $text[$rand.Next($text.Length)] })
}

function local:DivideFile([string]$SrcPath, [int]$PartSizeMB) {
    $FileData = [System.IO.File]::ReadAllBytes($SrcPath)
    $FileSize = $FileData.Length
    $PartSize = $PartSizeMB * 1024 * 1024
    $PartNum  = 0
    if ($FileSize -lt $PartSize) {
        while ($FileSize -gt 0) {
            $PartNum += 1
            $dname = [System.IO.Path]::GetDirectoryName($SrcPath)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($SrcPath)
            $PartPath = [System.IO.Path]::Combine($dname, $fname + ".div.$("{0:D3}" -f $PartNum)")
            $PartSize = [math]::Min($PartSize, $FileSize)
            $PartData = $FileData[0..($PartSize - 1)]
            [System.IO.File]::WriteAllBytes($PartPath, $PartData)
            $FileData = $FileData[$PartSize..($FileSize - 1)]
            $FileSize -= $PartSize
        }
    }
    return $PartNum
}

function local:ManipArchive($Path) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 本体処理
        if ((isArchive $Path) -or (isDividedArchive $Path)) {
            # 展開処理
            $dname = [System.IO.Path]::GetDirectoryName($Path)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($Path)
            $ename = [System.IO.Path]::GetExtension($Path)
            $sExtSrcPath = $Path
            $sExtDstPath = [System.IO.Path]::Combine($dname, $fname)
            if (isDividedArchive $sExtSrcPath) {
                if ($ename -ne ".001") {
                    Write-Host "$($Path)は分割圧縮ファイルの先頭ではありません"
                    return
                }
            }
            ExtArc7Z -SrcPath $sExtSrcPath -DstPath $sExtDstPath -DelSrc $false
        } else {
            # 圧縮処理
            $dname = [System.IO.Path]::GetDirectoryName($Path)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($Path)
            $sCmpSrcPath = $Path
            $sCmpDstPath = [System.IO.Path]::Combine($dname, $fname + ".zip")

            $ZipPwd = ""
            if ($Conf.Encrypt -eq $true) {
                $ZipPwd = GetRandomPassword
            }
            switch ($Conf.DivideType) {
                "None" { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
                "New"  { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd -DivideSize $Conf.DivideSize -DelSrc $false }
                "Old"  { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
            }
            $dname = [System.IO.Path]::GetDirectoryName($sCmpDstPath)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($sCmpDstPath)
            $sCmpBatPath = [System.IO.Path]::Combine($dname, $fname + ".bat")
            $sCmpPwdPath = [System.IO.Path]::Combine($dname, $fname + ".txt")
            if (($Conf.DivideType -eq [enmDivideType]::Old) -and ($Conf.DivideSize -gt 0)) {
                ## 旧式分割圧縮
                if (DivideFile $sCmpDstPath $Conf.DivideSize) {
                    Remove-Item $sCmpDstPath -Force
                    $sCmd = ""
                    $sCmd += "@echo off`r`n"
                    $sCmd += "copy /b ""$($fname).div.*"" ""$($fname).zip"""
                    Set-Content -LiteralPath $sCmpBatPath -Value $sCmd -Encoding "OEM"
                }
            }
            if ($Conf.Encrypt -eq $true) {
                Set-Clipboard $ZipPwd
                Set-Content -Path $sCmpPwdPath -Value $ZipPwd -Encoding "OEM"
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.zip")

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
    # 引数確認
    if ($args.Count -eq 0) {
        EditConfFile $Title $ConfPath
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            ManipArchive $arg
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
