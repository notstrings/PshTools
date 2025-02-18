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
    scoop install imagemagick
    scoop install ghostscript
    scoop install qpdf
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
    Enum enmResizeMode {
        S = 0
        M = 1
        L = 2
    }
    Enum enmPDFPaperSize {
        A1 = 0
        A2 = 1
        A3 = 2
        A4 = 3
        A5 = 4
    }
    class ManipImageConf {
        [System.ComponentModel.Category("Resize")]
        [System.ComponentModel.Description("S:640/M:800/L:1280")]
        [enmResizeMode]   `$ResizeMode

        [System.ComponentModel.Category("AnnotateTitle")]
        [bool]            `$AnnotateTitle
        [System.ComponentModel.Category("AnnotateTitle")]
        [System.ComponentModel.Description("%DN%=ディレクトリ名/%DN%=ファイル名")]
        [string]          `$AnnotateTitleText
        [System.ComponentModel.Category("AnnotateTitle")]
        [enmGravityType]  `$AnnotateTitlePos
        [System.ComponentModel.Category("AnnotateTitle")]
        [int]             `$AnnotateTitleSize
        [System.ComponentModel.Category("AnnotateTitle")]
        [System.ComponentModel.Description("#RRGGBB")]
        [string]          `$AnnotateTitleColor

        [System.ComponentModel.Category("AnnotateDetail")]
        [bool]            `$AnnotateDetail
        [System.ComponentModel.Category("AnnotateDetail")]
        [System.ComponentModel.Description("%DN%=ディレクトリ名/%DN%=ファイル名")]
        [string]          `$AnnotateDetailText
        [System.ComponentModel.Category("AnnotateDetail")]
        [enmGravityType]  `$AnnotateDetailPos
        [System.ComponentModel.Category("AnnotateDetail")]
        [int]             `$AnnotateDetailSize
        [System.ComponentModel.Category("AnnotateDetail")]
        [System.ComponentModel.Description("#RRGGBB")]
        [string]          `$AnnotateDetailColor

        [System.ComponentModel.Category("AnnotateBorder")]
        [bool]            `$AnnotateBorder
        [System.ComponentModel.Category("AnnotateBorder")]
        [int]             `$AnnotateBorderSize
        [System.ComponentModel.Category("AnnotateBorder")]
        [System.ComponentModel.Description("#RRGGBB")]
        [string]          `$AnnotateBorderColor

        [System.ComponentModel.Category("PDFPaperSize")]
        [enmPDFPaperSize] `$PDFPaperSize
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $Conf = New-Object ManipImageConf -Property @{
            ResizeMode = [enmResizeMode]::M

            AnnotateTitle = $true
            AnnotateTitleText = "%DN%"
            AnnotateTitlePos = [enmGravityType]::North
            AnnotateTitleSize = 50
            AnnotateTitleColor = "#000000"

            AnnotateDetail = $true
            AnnotateDetailText = "%FN%"
            AnnotateDetailPos = [enmGravityType]::South
            AnnotateDetailSize = 40
            AnnotateDetailColor = "#000000"

            AnnotateBorder = $true
            AnnotateBorderSize = 8
            AnnotateBorderColor = "#000000"

            PDFPaperSize = [enmPDFPaperSize]::a4
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [ManipImageConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([ManipImageConf]) $json
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

# ファイル・フォルダ名の処理
function local:ManipImage([string[]] $TargetPaths, [string] $Mode) {
    try {
        # 設定取得
        $Conf = LoadConfFile $ConfPath
        # 変換実施
        $IMPath = "magick.exe"
        $GSPath = "gswin64c.exe"
        $QPPath = "qpdf.exe"
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
                switch ($Conf.ResizeMode) {
                    "S" { $ResizeMode = "640x640>"   }
                    "M" { $ResizeMode = "800x800>"   }
                    "L" { $ResizeMode = "1280x1280>" }
                }
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
                        $arg += " -resize $ResizeMode"
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
                        $text0 = $Conf.AnnotateTitleText
                        $text0 = $text0 -replace "%DN%", ([System.IO.Path]::GetFileName($dname))
                        $text0 = $text0 -replace "%FN%", ([System.IO.Path]::GetFileNameWithoutExtension($srcpath))
                        $text1 = $Conf.AnnotateDetailText
                        $text1 = $text1 -replace "%DN%", ([System.IO.Path]::GetFileName($dname))
                        $text1 = $text1 -replace "%FN%", ([System.IO.Path]::GetFileNameWithoutExtension($srcpath))
                        $arg = ""
                        $arg += "convert"
                        $arg += " ""$srcpath"""
                        $arg += " -resize 1280x1280"                                    # リサイズ※フォントサイズ指定に苦労したくないから固定で適用する
                        $arg += " ( +clone -alpha opaque -fill white -colorize 100% )"  # 透過背景対策
                        $arg += " +swap -compose Over -composite"                       # 透過背景対策
                        if ($Conf.AnnotateTitle -eq $true -and $Conf.AnnotateTitleText -ne "") {
                            # 縁取りのために二度打ち
                            $arg += " -gravity $($Conf.AnnotateTitlePos.ToString())"                                                                                        # 上側テキスト合成
                            $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize $($Conf.AnnotateTitleSize)"                                                               # 上側テキスト合成
                            $arg += " ( -background none -stroke ""#FFFFFF""                     -strokewidth 2 -fill ""$($Conf.AnnotateTitleColor)"" label:""$text0"" )"   # 上側テキスト合成
                            $arg += " -composite"                                                                                                                           # 上側テキスト合成
                            $arg += " -gravity $($Conf.AnnotateTitlePos.ToString())"                                                                                        # 上側テキスト合成
                            $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize $($Conf.AnnotateTitleSize)"                                                               # 上側テキスト合成
                            $arg += " ( -background none -stroke ""$($Conf.AnnotateTitleColor)"" -strokewidth 0 -fill ""$($Conf.AnnotateTitleColor)"" label:""$text0"" )"   # 上側テキスト合成
                            $arg += " -composite"                                                                                                                           # 上側テキスト合成
                        }
                        if ($Conf.AnnotateDetail -eq $true -and $Conf.AnnotateDetailText -ne "") {
                            # 縁取りのために二度打ち
                            $arg += " -gravity $($Conf.AnnotateDetailPos.ToString())"                                                                                       # 下側テキスト合成
                            $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize $($Conf.AnnotateDetailSize)"                                                              # 下側テキスト合成
                            $arg += " ( -background none -stroke ""#FFFFFF""                      -strokewidth 2 -fill ""$($Conf.AnnotateDetailColor)"" label:""$text1"" )" # 下側テキスト合成
                            $arg += " -composite"                                                                                                                           # 下側テキスト合成
                            $arg += " -gravity $($Conf.AnnotateDetailPos.ToString())"                                                                                       # 下側テキスト合成
                            $arg += " -font ""MS-Mincho-&-MS-PMincho"" -pointsize $($Conf.AnnotateDetailSize)"                                                              # 下側テキスト合成
                            $arg += " ( -background none -stroke ""$($Conf.AnnotateDetailColor)"" -strokewidth 0 -fill ""$($Conf.AnnotateDetailColor)"" label:""$text1"" )" # 下側テキスト合成
                            $arg += " -composite"                                                                                                                           # 下側テキスト合成
                        }
                        if ($Conf.AnnotateBorder -eq $true) {
                            $arg += " -bordercolor ""$($Conf.AnnotateBorderColor)"""                        # 枠線色
                            $arg += " -border ""$($Conf.AnnotateBorderSize)x$($Conf.AnnotateBorderSize)"""  # 枠線
                        }
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
                # $srcpath = $TargetPath
                $dstpath = [System.IO.Path]::Combine($dname, "Conv", "Combined.pdf")
                $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                $arg = ""
                $arg += " -q -dNOPAUSE -dBATCH -sDEVICE=pdfwrite"
                $arg += " -sOutputFile=""$dstpath"" $CombineFiles"
                $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                ## 後始末
                Remove-Item -LiteralPath ([System.IO.Path]::Combine($dname, "Conv", "Temp")) -Recurse -Force
            }
            "PDF平坦化" {
                # PDF中のスタンプなどを平坦化しAcrobatReaderでは簡単に編集できなくする
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += " --flatten-annotations=all"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$QPPath""" -ArgumentList $arg
                    }
                }
            }
            "PDF圧縮" {
                # PDF内部の画像等を圧縮する
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
            "PDF用紙リサイズ" {
                # シートサイズが不均等なExcelをPDF出力した場合対策
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -eq ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += " -sDEVICE=pdfwrite -dCompatibilityLevel=1.4"                                           # PDF1.4互換出力
                        $arg += " -sPAPERSIZE=$($Conf.PDFPaperSize.ToString().ToLower()) -dFIXEDMEDIA -dPDFFitPage"     # 用紙サイズ指定
                        $arg += " -dNOPAUSE -dBATCH -dQUIET"                                                            # ごにょごにょ
                        $arg += " -sOutputFile=""$dstpath"" ""$srcpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$GSPath""" -ArgumentList $arg
                    }
                }
            }
            "PDF捨印生成" {
                # PDF捨印作成用
                foreach ($TargetPath in $TargetPaths) {
                    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
                    $ename = [System.IO.Path]::GetExtension($TargetPath)
                    if ($ename.ToLower() -ne ".pdf") {
                        $srcpath = $TargetPath
                        $dstpath = [System.IO.Path]::Combine($dname, "Conv", $fname + ".pdf")
                        $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                        $arg = ""
                        $arg += "convert"
                        $arg += " -resize 400x400 -density 1024"
                        $arg += " ""$srcpath"" ""$dstpath"""
                        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList $arg
                    }
                }
            }
            Default {
                EditConfFile $Title $ConfPath
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.png")

try {
    $null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
    # 本体処理
    $ret = ShowFileListDialogWithOption `
            -Title $Title `
            -Message "対象画像ファイルをドラッグ＆ドロップしてください`n入力可能形式はbmp/jpg/jpeg/gif/tif/tiff/png/svg/pdfです`n※PDFは扱いが特殊ですんで一部無視されたりします" `
            -FileList $args `
            -FileFilter "\.(bmp|jpg|jpeg|gif|tif|tiff|png|svg|pdf)$" `
            -Options @("PNG変換", "リサイズ", "トリミング", "傾き補正", "注釈付記", 
                       "PDF変換個別", "PDF変換統合", "PDF平坦化", "PDF圧縮", "PDF用紙リサイズ", "PDF捨印生成", 
                       "設定編集")
    if ($ret[0] -eq "OK") {
        ManipImage $ret[1] $ret[3]
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
