$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

# ファイル・フォルダ名の処理
function local:ExecImageManip([string] $TargetPath, [string] $Mode) {
    try {
        $IMPath = "C:\ImageMagick\magick.exe"
        $GSPath = "C:\ImageMagick\CommonFiles\Ghostscript\bin\gswin64c"
        $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
        $ename = [System.IO.Path]::GetExtension($TargetPath).ToLower()
        $srcpath = $TargetPath
        $dstpath = [System.IO.Path]::Combine($dname, "Conv")
        $null = New-Item $dstpath -ItemType Directory -ErrorAction SilentlyContinue
        # 変換実施
        switch ($Mode) {
            "CONVERT PNG※" {
                if ($ename -ne ".pdf") {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + ".png")
                    & $IMPath convert $srcpath $dstpath
                } else {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname)
                    $null = New-Item $dstpath -ItemType Directory -ErrorAction SilentlyContinue
                    $dstpath = [System.IO.Path]::Combine($dstpath, "%04d.png")
                    & $IMPath convert -density 300 -alpha remove $srcpath $dstpath
                }
            }
            "RESIZE" {
                if ($ename -ne ".pdf") {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + $ename)
                    & $IMPath convert -resize "800x800>" $srcpath $dstpath
                }
            }
            "TRIM" {
                if ($ename -ne ".pdf") {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + $ename)
                    & $IMPath convert -fuzz 10% -trim $srcpath $dstpath
                }
            }
            "DESKEW" {
                if ($ename -ne ".pdf") {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + $ename)
                    & $IMPath convert -deskew 40% $srcpath $dstpath
                }
            }
            "ANOTATE" {
                if ($ename -ne ".pdf") {
                    $text0 = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($srcpath))
                    $text1 = [System.IO.Path]::GetFileNameWithoutExtension($srcpath)
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + $ename)
                    & $IMPath convert  `
                        $srcpath `
                        `( +clone -alpha opaque -fill white -colorize 100% `) `
                        +swap -compose Over -composite `
                        -gravity north `
                        -font "MS-Mincho-&-MS-PMincho" -pointsize 35 `
                        `( -background none -fill "#000000" label:"$text0" `) `
                        -composite `
                        -gravity south `
                        -font "MS-Mincho-&-MS-PMincho" -pointsize 25 `
                        `( -background none -fill "#000000" label:"$text1" `) `
                        -composite `
                        -bordercolor "#000000" -border "8x8" `
                        $dstpath
                }
            }
            "CONVERT PDF" {
                if ($ename -ne ".pdf") {
                    # スタンプサイズ
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + ".pdf")
                    & $IMPath convert -resize "x400" -density 1024 $srcpath $dstpath
                }
            }
            "COMPRESS PDF※" {
                if ($ename -eq ".pdf") {
                    $dstpath = [System.IO.Path]::Combine($dstpath, $fname + ".pdf")
                    & $GSPath `
                        -sDEVICE=pdfwrite -dCompatibilityLevel="1.4" `
                        -dPDFSETTINGS=/screen `
                        -dNOPAUSE -dBATCH -dQUIET -dPDFFitPage `
                        -sOutputFile="$dstpath" "$srcpath"
                }
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

try {
    $null = Write-Host "---ManipImage---"
    $ret = ShowFileListDialogWithOption `
            -Title "画像操作" `
            -Message "対象画像ファイルをドラッグ＆ドロップしてください`n入力可能形式はbmp/jpg/jpeg/gif/tif/tiff/png/svg/pdfです`n※pdfは一部処理のみ有効" `
            -FileList $args `
            -FileFilter "\.(bmp|jpg|jpeg|gif|tif|tiff|png|svg|pdf)$" `
            -Options @("CONVERT PNG※", "RESIZE", "TRIM", "DESKEW", "ANOTATE", "CONVERT PDF", "COMPRESS PDF※")
    if ($ret[0] -eq "OK") {
        foreach ($elm in $ret[1]) {
            if (Test-Path -LiteralPath $elm) {
                ExecImageManip $elm $ret[2]
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
 