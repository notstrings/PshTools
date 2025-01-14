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

Enum enmDivideType {
    None = 0
    New  = 1
    Old  = 2
}
class Conf {
    [bool]           $Encrypt
    [enmDivideType]  $DivideType
    [int]            $DivideSize
}

# 設定初期化
function local:InitConf([string] $sPath) {
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $conf = New-Object Conf -Property @{
            Encrypt = $false
            DivideType = [enmDivideType]::None
            DivideSize = 1
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

# アーカイブファイル判定
function isArchive($sPath) {
    return (([System.IO.Path]::GetExtension($sPath).ToUpper()) -in @(".7Z", ".GZ", ".ZIP", ".BZ2", ".TAR", ".LZH", ".LZS", ".LHA", ".GZIP", ".LZMA"))
}

# 分割圧縮ファイル判定
function local:isDividedArchive($sPath) {
    return (([System.IO.Path]::GetExtension($sPath).ToUpper()) -match "\.\d{3}$")
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

function local:ManipArchive($sPath) {
    if ((isArchive $sPath) -or (isDividedArchive $sPath)) {
        # 展開処理
        $dname = [System.IO.Path]::GetDirectoryName($sPath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($sPath)
        $ename = [System.IO.Path]::GetExtension($sPath)
        $sExtSrcPath = $sPath
        $sExtDstPath = [System.IO.Path]::Combine($dname, $fname)
        if (isDividedArchive $sExtSrcPath) {
            if ($ename -ne ".001") {
                Write-Host "$($sPath)は分割圧縮ファイルの先頭ではありません"
                return
            }
        }
        ExtArc7Z -SrcPath $sExtSrcPath -DstPath $sExtDstPath -DelSrc $false
    } else {
        # 圧縮処理
        $dname = [System.IO.Path]::GetDirectoryName($sPath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($sPath)
        $sCmpSrcPath = $sPath
        $sCmpDstPath = [System.IO.Path]::Combine($dname, $fname + ".zip")

        $ZipPwd = ""
        if ($conf.Encrypt -eq $true) {
            $ZipPwd = GetRandomPassword
        }
        switch ($conf.DivideType) {
            "None" { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
            "New"  { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd -DivideSize $conf.DivideSize -DelSrc $false }
            "Old"  { $sCmpDstPath = CmpArc7Z -SrcPath $sCmpSrcPath -DstPath $sCmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
        }
        $dname = [System.IO.Path]::GetDirectoryName($sCmpDstPath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($sCmpDstPath)
        $sCmpBatPath = [System.IO.Path]::Combine($dname, $fname + ".bat")
        $sCmpPwdPath = [System.IO.Path]::Combine($dname, $fname + ".txt")
        if (($conf.DivideType -eq [enmDivideType]::Old) -and ($conf.DivideSize -gt 0)) {
            ## 旧式分割圧縮
            if (DivideFile $sCmpDstPath $conf.DivideSize) {
                Remove-Item $sCmpDstPath -Force
                $sCmd = ""
                $sCmd += "@echo off`r`n"
                $sCmd += "copy /b ""$($fname).div.*"" ""$($fname).zip"""
                Set-Content -LiteralPath $sCmpBatPath -Value $sCmd -Encoding "OEM"
            }
        }
        if ($conf.Encrypt -eq $true) {
            Set-Clipboard $ZipPwd
            Set-Content -Path $sCmpPwdPath -Value $ZipPwd -Encoding "OEM"
        }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.zip")

try {
    $null = Write-Host "---$Title---"
    # 設定取得
    InitConf $ConfPath
    $conf = LoadConf $ConfPath
    # 引数確認
    if ($args.Count -eq 0) {
        EditConf $ConfPath
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
