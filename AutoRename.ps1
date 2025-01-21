$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    class AutoRenameConf {
        [bool] `$RestrictZen
        [bool] `$RestrictHan
        [bool] `$RestrictDate
        [string] `$RestrictDateFormat
        [bool] `$RestrictBlank
        [bool] `$RemoveBracket
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object AutoRenameConf -Property @{
            RestrictZen = $true
            RestrictHan = $true
            RestrictDate = $true
            RestrictDateFormat = "yyyyMMdd"
            RestrictBlank = $true
            RemoveBracket = $false
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [AutoRenameConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([AutoRenameConf]) $json
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
function local:AutoRenameFile([System.IO.FileInfo] $Target) {
    AutoRename $Target.FullName $Target.LastWriteTime $false
}

# フォルダ
function local:AutoRenameDir([System.IO.DirectoryInfo] $Target) {
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        AutoRenameDir  $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        AutoRenameFile $elm
    }
    AutoRename $Target.FullName $Target.CreationTime $true
}

# ファイル・フォルダ名の処理
function local:AutoRename([string] $TargetPath, [datetime] $TargetDate, [bool] $isDir) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 修正前名称
        $srcpath = $TargetPath
        # 修正後名称
        $dstpath = $TargetPath
        if ($isDir -eq $false) {
            $dname = [System.IO.Path]::GetDirectoryName($dstpath)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($dstpath)
            $ename = [System.IO.Path]::GetExtension($dstpath)
        } else {
            $dname = [System.IO.Path]::GetDirectoryName($dstpath)
            $fname = [System.IO.Path]::GetFileName($dstpath)
            $ename = ""
        }
        if ($Conf.RestrictZen -eq $true) {
            $fname = RestrictTextZen   -Text $fname -Chars "Ａ-Ｚａ-ｚ０-９　（）［］｛｝"
        }
        if ($Conf.RestrictHan -eq $true) {
            $fname = RestrictTextHan   -Text $fname
        }
        if ($Conf.RestrictDate -eq $true) {
            $fname = RestrictTextDate  -Text $fname -Format $Conf.RestrictDateFormat -RefDate $TargetDate
        }
        if ($Conf.RestrictBlank -eq $true) {
            $fname = RestrictTextBlank -Text $fname
        }
        if ($Conf.RemoveBracket -eq $true) {
            $fname = RemoveAllBrackets -Text $fname # ファイル名が重複すると自分で括弧付けるんだが、まぁあれば便利
        }
        $dstpath = [System.IO.Path]::Combine($dname, $fname + $ename)
        # 必要があればリネーム
        if ($fname -ne "") {
            if ($srcpath -ne $dstpath) {
                $uniq = GenUniqName $dstpath $isDir
                $null = Write-Host "---"
                $null = Write-Host "src : $srcpath"
                $null = Write-Host "dst : $uniq"
                $null = MoveItemWithProgress $srcpath $uniq
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
    if ($args.Count -eq 0) {
        EditConfFile $Title $ConfPath
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                AutoRenameDir  (Get-Item $arg)
            } else {
                AutoRenameFile (Get-Item $arg)
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
