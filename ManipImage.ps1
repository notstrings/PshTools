$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

# セットアップ
function local:Setup() {
    scoop install main/imagemagick
    scoop install main/ghostscript
}

# ファイル・フォルダ名の処理
function local:ManipImage([string[]] $TargetPaths, [string] $Mode) {
    try {
        $IMPath = "magick.exe"
        $GSPath = "gswin64c.exe"
        # 変換実施
        # ・クソ長いが分割した所で結局大した意味もない
        switch ($Mode) {
            "PNG変換" {
                # 入力ファイルをPNGに変換する
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".png")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " ""$srcpath"" ""$dstpath"" "
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    } else {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname, "%04d.png")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " -density 300 -alpha remove"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "リサイズ" {
                # 入力ファイルをリサイズする
                # ・サイズはどうせほぼ固定だろうからハードコーディング
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + $ename)
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " -resize 800x800>"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "トリミング" {
                # 入力ファイルの余白をトリミングする
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + $ename)
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " -fuzz 10% -trim"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "傾き補正" {
                # 入力ファイルに傾き補正を掛ける
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + $ename)
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " -deskew 40%"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "注釈付記" {
                # 入力ファイルに注釈を書き込んで枠を付ける
                # ・上側注記がフォルダ名で下側注記がファイル名
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + $ename)
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $text0 = [System.IO.Path]::GetFileName($dname)
                        $text1 = [System.IO.Path]::GetFileNameWithoutExtension($srcpath)
                        $arg = ""
                        $arg += "convert"
                        $arg += " ""$srcpath"""
                        $arg += " -resize 800x800>"                                             # リサイズ
                        $arg += " ( +clone -alpha opaque -fill white -colorize 100% )"          # 透過背景対策
                        $arg += " +swap -compose Over -composite"                               # 透過背景対策
                        $arg += " -gravity north"                                               # 上側テキスト合成
                        $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize 35"               # 上側テキスト合成
                        $arg += " ( -background none -fill ""#000000"" label:""$text0"" )"      # 上側テキスト合成
                        $arg += " -composite"                                                   # 上側テキスト合成
                        $arg += " -gravity south"                                               # 下側テキスト合成
                        $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize 25"               # 下側テキスト合成
                        $arg += " ( -background none -fill ""#000000"" label:""$text1"" )"      # 下側テキスト合成
                        $arg += " -composite"                                                   # 下側テキスト合成
                        $arg += " -bordercolor ""#000000"" -border ""8x8"""                     # 枠線付与
                        $arg += " ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "PDF変換個別" {
                # 入力ファイルを個々にPDFにする
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        # PDFは何もしないでコピー
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = Copy-Item -LiteralPath $srcpath -Destination $dstpath -Force
                    } else {
                        # PDF以外は形式変換を実施
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            "PDF変換統合" {
                # 入力ファイルを単一のPDFにする
                ## PDF変換
                $CombineFiles = ""
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        # PDFは何もしないでコピー
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", "Temp", $fname + ".pdf")
                        $null = Copy-Item -LiteralPath $srcpath -Destination $dstpath -Force
                    } else {
                        # PDF以外は形式変換を実施
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", "Temp", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                    $CombineFiles += " ""$dstpath"""
                }
                ## PDF統合
                $srcpath = $TargetPath
                $dstpath = [System.IO.Path]::Combine($dname, "Conv", "Combined.pdf")
                $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                $arg = ""
                $arg += " -q -dNOPAUSE -dBATCH -sDEVICE=pdfwrite"
                $arg += " -sOutputFile=""$dstpath"" $CombineFiles"
                $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                ## 後始末
                Remove-Item -LiteralPath ([System.IO.Path]::Combine($dname, "Conv", "Temp")) -Recurse -Force
            }
            "PDF圧縮" {
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += " -sDEVICE=pdfwrite -dCompatibilityLevel=1.4"   # PDF1.4互換出力
                        $arg += " -dPDFSETTINGS=/screen -dPDFFitPage"           # スクリーンレベルの圧縮率
                        $arg += " -dNOPAUSE -dBATCH -dQUIET"                    # ごにょごにょ
                        $arg += " -sOutputFile=""$dstpath"" ""$srcpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                    }
                }
            }
            "PDFリサイズA3" {
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += " -sDEVICE=pdfwrite -dCompatibilityLevel=1.4"   # PDF1.4互換出力
                        $arg += " -sPAPERSIZE=a3 -dFIXEDMEDIA -dPDFFitPage"     # A3を指定
                        $arg += " -dNOPAUSE -dBATCH -dQUIET"                    # ごにょごにょ
                        $arg += " -sOutputFile=""$dstpath"" ""$srcpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                    }
                }
            }
            "PDFリサイズA4" {
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += " -sDEVICE=pdfwrite -dCompatibilityLevel=1.4"   # PDF1.4互換出力
                        $arg += " -sPAPERSIZE=a4 -dFIXEDMEDIA -dPDFFitPage"     # A4を指定
                        $arg += " -dNOPAUSE -dBATCH -dQUIET"                    # ごにょごにょ
                        $arg += " -sOutputFile=""$dstpath"" ""$srcpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                    }
                }
            }
            "PDF捨印生成" {
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        # スタンプ生成用にサイズ調整が必要になる
                        $arg = ""
                        $arg += "convert"
                        $arg += " -resize x400 -density 1024"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---ManipImage---"
    $ret = ShowFileListDialogWithOption `
            -Title "画像操作" `
            -Message "対象画像ファイルをドラッグ＆ドロップしてください`n入力可能形式はbmp/jpg/jpeg/gif/tif/tiff/png/svg/pdfです`n※PDFは扱いが特殊ですんで一部無視されたりします" `
            -FileList $args `
            -FileFilter "\.(bmp|jpg|jpeg|gif|tif|tiff|png|svg|pdf)$" `
            -Options @("PNG変換", "リサイズ", "トリミング", "傾き補正", "注釈付記", "PDF変換個別",
                       "PDF変換統合", "PDF圧縮", "PDFリサイズA3", "PDFリサイズA4", "PDF捨印生成")
    if ($ret[0] -eq "OK") {
        ManipImage $ret[1] $ret[2]
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
