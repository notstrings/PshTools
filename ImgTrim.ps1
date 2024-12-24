$ErrorActionPreference = "Stop"

$ExePath = "C:\ImageMagick\magick.exe"
# imagemagickインストール
# portableGhostscriptをimagemagickの下CommonFilesに放り込む
# imagemagickのdelegate.xlsの@PSDelegate@を.\CommonFilesにリプレース

# ファイル・フォルダ名の処理
function local:ExecImageMagick([string] $TargetPath) {
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
    $ret = ShowFileListDialogWithOption `
            -Title "ImageMagickHelper" `
            -Message "対象画像ファイルをD&Dしてください" `
            -FileList $args `
            -FileFilter "\.(bmp|jpg|jpeg|gif|tif|tiff|png)$" `
            -Options @("RESIZE", "TRIM", "ANNOTATE", "COMPARE")
    if ($ret[0] -eq "OK") {
        foreach ($elm in $ret[1]) {
            if (Test-Path -LiteralPath $elm) {
                ExecImageMagick $elm
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
 