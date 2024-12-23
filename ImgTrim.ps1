$ErrorActionPreference = "Stop"

$ExePath = "C:\ImageMagick\magick.exe"
# imagemagickインストール
# portableGhostscriptをimagemagickの下CommonFilesに放り込む
# imagemagickのdelegate.xlsの@PSDelegate@を.\CommonFilesにリプレース

# ファイル名の処理
function local:ProcessFPath([System.IO.FileInfo] $Target) {
    ProcessNode $Target.FullName
}

# フォルダ名の処理
function local:ProcessDPath([System.IO.DirectoryInfo] $Target) {
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ProcessDPath $elm
    }
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ProcessFPath $elm
    }
}

# ファイル・フォルダ名の処理
function local:ProcessNode([string] $TargetPath) {
    try {
        # 入力ファイル名
        $srcpath = $TargetPath

        # 出力ファイル名
        $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
        $ename = ".png"
        $dstpath = [System.IO.Path]::Combine($dname, "out", $fname + $ename)

        # 変換実施
        & $ExePath convert """$srcpath""" -fuzz 10% -resize "800x800>" """$dstpath"""
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

try {
    if ($args.Length -eq 0) {
        exit 1
    }
    $null = Write-Host "<<Start>>"
    ForEach ($arg in $args) {
        if( Test-Path -LiteralPath $arg ){
            if ((Get-Item $arg).PSIsContainer) {
                ProcessDPath (Get-Item $arg)
            } else {
                ProcessFPath (Get-Item $arg)
            }
        }
    }
    $null = Write-Host "<<End>>"
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
