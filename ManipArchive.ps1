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

function local:DivideFile([string]$SrcPath, [string]$DstPath, [int]$sizeMB) {
    $cnt  = 0
    $PartSize = $sizeMB * 1024 * 1024
    $PartBuff = New-Object byte[] $PartSize
    $FileStrm = [System.IO.File]::OpenRead($SrcPath)
    try {
        while ($FileStrm.Position -lt $FileStrm.Length) {
            $cnt += 1
            $PartPath = [System.IO.Path]::Combine($DstPath, "$([System.IO.Path]::GetFileName($SrcPath)).$($cnt.ToString("000"))")
            $PartStrm = [System.IO.File]::Create($PartPath)
            try {
                $PartRead = $FileStrm.Read($PartBuff, 0, $PartSize)
                $PartStrm.Write($PartBuff, 0, $PartRead)
            } finally {
                $PartStrm.Close()
            }
        }
    } finally {
        $FileStrm.Close()
    }
    return $cnt
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
            $ExtSrcPath = $Path
            $ExtDstPath = [System.IO.Path]::Combine($dname, $fname)
            if (isDividedArchive $ExtSrcPath) {
                if ($ename -ne ".001") {
                    Write-Host "$($Path)は分割圧縮ファイルの先頭ではありません"
                    return
                }
            }
            ExtArc7Z -SrcPath $ExtSrcPath -DstPath $ExtDstPath -DelSrc $false
        } else {
            # 圧縮処理
            $dname = [System.IO.Path]::GetDirectoryName($Path)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($Path)
            $CmpSrcPath = $Path
            $CmpDstPath = [System.IO.Path]::Combine($dname, $fname + ".zip")

            $ZipPwd = ""
            if ($Conf.Encrypt -eq $true) {
                $ZipPwd = GetRandomPassword
            }
            switch ($Conf.DivideType) {
                "None" { $CmpDstPath = CmpArc7Z -SrcPath $CmpSrcPath -DstPath $CmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
                "New"  { $CmpDstPath = CmpArc7Z -SrcPath $CmpSrcPath -DstPath $CmpDstPath -ZipPwd $ZipPwd -DivideSize $Conf.DivideSize -DelSrc $false }
                "Old"  { $CmpDstPath = CmpArc7Z -SrcPath $CmpSrcPath -DstPath $CmpDstPath -ZipPwd $ZipPwd                              -DelSrc $false }
            }
            if (($Conf.DivideType -eq [enmDivideType]::Old) -and ($Conf.DivideSize -gt 0)) {
                # 旧式分割圧縮
                $dname = [System.IO.Path]::GetDirectoryName($CmpDstPath)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($CmpDstPath)
                $ename = [System.IO.Path]::GetExtension($CmpDstPath)
                $CmpDivPath = [System.IO.Path]::Combine($dname, "$($fname + $ename).div")
                $CmpBatPath = [System.IO.Path]::Combine($CmpDivPath, "$($fname + $ename).bat")
                $null = New-Item $CmpDivPath -ItemType Directory -ErrorAction SilentlyContinue
                $DivNum = DivideFile $CmpDstPath $CmpDivPath $Conf.DivideSize
                if ($DivNum -gt 0) {
                    Remove-Item $CmpDstPath -Force
                    $sCmd = ""
                    $sCmd += "@echo off`r`n"
                    $sCmd += "copy /b ""$($fname).*"" ""$($fname + $ename)"""
                    Set-Content -LiteralPath $CmpBatPath -Value $sCmd -Encoding "OEM"
                }
            }
            if ($Conf.Encrypt -eq $true) {
                # 暗号化パスワード
                $dname = [System.IO.Path]::GetDirectoryName($CmpDstPath)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($CmpDstPath)
                $ename = [System.IO.Path]::GetExtension($CmpDstPath)
                $CmpPwdPath = [System.IO.Path]::Combine($dname, $fname + ".txt")
                Set-Clipboard $ZipPwd
                Set-Content -Path $CmpPwdPath -Value $ZipPwd -Encoding "OEM"
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー２")

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
