param (
    [Parameter(Mandatory = $true)]  [string] $sSrcPath,
    [Parameter(Mandatory = $true)]  [string] $sDstPath,
    [Parameter(Mandatory = $false)] [long]   $lTgtPageStt = 1,
    [Parameter(Mandatory = $false)] [long]   $lTgtPageEnd = 0,
    [Parameter(Mandatory = $true)]  [string] $sTopTxt,
    [Parameter(Mandatory = $true)]  [string] $sMdlTxt,
    [Parameter(Mandatory = $true)]  [string] $sBtmTxt,
    [Parameter(Mandatory = $false)] [double] $dStumpPosX = 50,
    [Parameter(Mandatory = $false)] [double] $dStumpPosY = 50,
    [Parameter(Mandatory = $false)] [double] $dStumpPosSZ = 10
)

# セットアップ
function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install nuget
    scoop install gh

    # セットアップ先
    $ToolDir = [System.IO.Path]::Combine($PSScriptRoot, "tool")
    $null = New-Item $ToolDir -ItemType Directory -ErrorAction SilentlyContinue

    # rsvg-convert
    $ToolPath01 = [System.IO.Path]::Combine($ToolDir, "rsvg-convert.x86_64.zip")
    if ((Test-Path $ToolPath01) -eq $false) {
        gh api -H 'Accept: application/octet-stream' /repos/miyako/console-rsvg-convert/contents/rsvg-convert.x86_64.zip > rsvg-convert.x86_64.zip
        tar.exe "-zvxf" $ToolPath01 -C $ToolDir
    }

    # PDFSharp
    $ToolPath02 = [System.IO.Path]::Combine($ToolDir, "PDFsharp.1.50.5147")
    if ((Test-Path $ToolPath02) -eq $false) {
        nuget install PdfSharp -Version 1.50.5147 -OutputDirectory $ToolDir -Source 'https://api.nuget.org/v3/index.json'
    }
}

<#
.SYNOPSIS
    日付印PDFを作成します
.PARAMETER sPath
    出力するPDFファイルのパスを指定します
.PARAMETER sTopTxt
    円の上部に表示するテキストを指定します
.PARAMETER sMdlTxt
    円の中央に表示するテキストを指定します
.PARAMETER sBtmTxt
    円の下部に表示するテキストを指定します
.EXAMPLE
    MakePDFDateStamp "$($ENV:USERPROFILE)\Desktop\stamp.pdf" "上段" "中段" "下段"
#>
function local:MakePDFDateStamp {
    param (
        [Parameter(Mandatory = $true)] [string] $sPath,
        [Parameter(Mandatory = $true)] [string] $sTopTxt,
        [Parameter(Mandatory = $true)] [string] $sMdlTxt,
        [Parameter(Mandatory = $true)] [string] $sBtmTxt
    )
    begin {}
    process {
        # 日付印SVG
        # ・一時ファイル作るのダルいので標準入力でSVGを放り込む
        # ・UTF-8文字列にしなければならない
        chcp 65001
        $text = ""
        $text += "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>"
        $text += "<!DOCTYPE svg PUBLIC ""-//W3C//DTD SVG 1.1//EN"" ""http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"">"
        $text += "<svg"
        $text += "    xmlns=""http://www.w3.org/2000/svg"""
        $text += "    xmlns:xlink=""http://www.w3.org/1999/xlink"""
        $text += "    width=""263px"""
        $text += "    height=""263px"""
        $text += ">"
        $text += "    <g opacity=""0.5"">"
        $text += "        <circle id=""circle"" cx=""131.50"" cy=""131.50"" r=""120"" style=""fill:none;stroke:red;stroke-width:3px""/>"
        $text += "        <text x=""131.5"" y=""85""  font-family=""Yu Gothic"" font-size=""54"" text-anchor=""middle"" style=""fill:red;"">"
        $text += "            $($sTopTxt)"
        $text += "        </text>"
        $text += "        <text x=""131.5"" y=""148"" font-family=""Yu Gothic"" font-size=""44"" text-anchor=""middle"" style=""fill:red;"">"
        $text += "            $($sMdlTxt)"
        $text += "        </text>"
        $text += "        <text x=""131.5"" y=""214"" font-family=""Yu Gothic"" font-size=""46"" text-anchor=""middle"" style=""fill:red;"">"
        $text += "            $($sBtmTxt)"
        $text += "        </text>"
        $text += "        <line id=""line"" x1=""18"" y1=""96""  x2=""246"" y2=""96""  style=""stroke:red;fill:none;stroke-width:2px""/>"
        $text += "        <line id=""line"" x1=""18"" y1=""166"" x2=""246"" y2=""166"" style=""stroke:red;fill:none;stroke-width:2px""/>"
        $text += "    </g>"
        $text += "</svg>"

        # SVGをPDFに変換
        # ・出力パスに日本語が含まれると駄目っぽいので一時フォルダに出力する
        $TmpPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName())
        $procInfo = New-Object System.Diagnostics.ProcessStartInfo
        $procInfo.FileName = [System.IO.Path]::Combine($PSScriptRoot, "tool", "rsvg-convert.exe")
        $procInfo.Arguments = "-f pdf -o " + $TmpPath
        $procInfo.WorkingDirectory = $PSScriptRoot
        $procInfo.UseShellExecute = $false
        $procInfo.RedirectStandardInput = $true
        $procInfo.RedirectStandardOutput = $false
        $proc = [System.Diagnostics.Process]::Start($procInfo)
        if ($proc) { 
            $proc.StandardInput.Write($text)
            $proc.StandardInput.Close()
            $proc.WaitForExit()
            $proc.Dispose()
            Move-Item -LiteralPath $TmpPath -Destination $sPath -Force
        }
    }
    end {}
}

<#
.SYNOPSIS
    PDFファイルに指定したPDFスタンプを捺印します
.PARAMETER sSrcPath
    スタンプを捺印されるPDFファイルのパスを指定します
.PARAMETER sDstPath
    スタンプ後の出力先PDFファイルのパスを指定します
.PARAMETER lTgtPageStt
    スタンプを追加する開始ページの番号を指定します
    ページ番号は1から始まり0以下は最終ページとして扱われます
.PARAMETER lTgtPageEnd
    スタンプを追加する終了ページの番号を指定します
    ページ番号は1から始まり0以下は最終ページとして扱われます
.PARAMETER sStumpPath
    スタンプPDFファイルのパスを指定します
.PARAMETER dStumpPosX
    スタンプを配置するX座標の割合を指定します
    0%は左端で100%は右端に位置します
.PARAMETER dStumpPosY
    スタンプを配置するY座標の割合を指定します
    0%は上端で100%は下端に位置します
.PARAMETER dStumpPosSZ
    スタンプのサイズの割合を指定します
    ページの短編を基準とした割合を指定します
.EXAMPLE
    AnnotatePDFStamp "$($ENV:USERPROFILE)\Desktop\src.pdf" "$($ENV:USERPROFILE)\Desktop\dst.pdf" 1 0 "$($ENV:USERPROFILE)\Desktop\stamp.pdf" 50 50 10
#>
function local:AnnotatePDFStamp {
    param (
        [Parameter(Mandatory = $true)]  [string] $sSrcPath,
        [Parameter(Mandatory = $true)]  [string] $sDstPath,
        [Parameter(Mandatory = $false)] [long]   $lTgtPageStt = 1,
        [Parameter(Mandatory = $false)] [long]   $lTgtPageEnd = 0,
        [Parameter(Mandatory = $true)]  [string] $sStumpPath,
        [Parameter(Mandatory = $false)] [double] $dStumpPosX = 50,
        [Parameter(Mandatory = $false)] [double] $dStumpPosY = 50,
        [Parameter(Mandatory = $false)] [double] $dStumpPosSZ = 10
    )
    begin {
        [System.Reflection.Assembly]::LoadFrom([System.IO.Path]::Combine($PSScriptRoot, "tool", "PDFsharp.1.50.5147", "lib", "net20", "PdfSharp.dll")) | Out-Null
    }
    process {
        # 入力ファイルを開く
        $pdfrdr = [PdfSharp.Pdf.IO.PdfReader]::Open([string]$sSrcPath, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
    
        # 制御対象ページを解釈
        # ・利便性のために最終ページを0ページ目として指定する事ができるようにしておく
        $pagecount = $pdfrdr.pagecount
        if ($lTgtPageStt -le 0){ $lTgtPageStt = $pagecount }
        if ($lTgtPageEnd -le 0){ $lTgtPageEnd = $pagecount }
        $lTgtPageStt = [System.Math]::Max($lTgtPageStt, 1)
        $lTgtPageStt = [System.Math]::Min($lTgtPageStt, $pagecount)
        $lTgtPageEnd = [System.Math]::Max($lTgtPageEnd, 1)
        $lTgtPageEnd = [System.Math]::Min($lTgtPageEnd, $pagecount)
    
        # 描画本体
        $pdfstm = [PdfSharp.Drawing.XPdfForm]::FromFile([string]$sStumpPath)
        $pdfdoc = New-Object PdfSharp.Pdf.PdfDocument
        for ($idx = 1; $idx -le $pagecount; $idx++) {
            $page = $pdfdoc.AddPage($pdfrdr.Pages[$idx - 1]) 
            if (($idx -ge $lTgtPageStt) -And ($idx -le $lTgtPageEnd)) {
                # 出力座標を当該ページレイアウトに合わせる
                $gfx = [PdfSharp.Drawing.XGraphics]::FromPdfPage($page)
                $rot = ((($page.Rotate % 360) + 360 ) % 360)
                switch ( $rot ) # 座標系
                {
                      0 { $gfx.RotateTransform(   0); $gfx.TranslateTransform(                  0,                  0); }
                     90 { $gfx.RotateTransform( -90); $gfx.TranslateTransform(-$page.Height.Point,                  0); }
                    180 { $gfx.RotateTransform(-180); $gfx.TranslateTransform( -$page.Width.Point,-$page.Height.Point); }
                    270 { $gfx.RotateTransform(-270); $gfx.TranslateTransform(                  0, -$page.Width.Point); }
                }
                switch ( $rot ) # 領域サイズ
                {
                      0 { $gwd=$page.Width.Point;  $ght=$page.Height.Point; }
                     90 { $gwd=$page.Height.Point; $ght=$page.Width.Point;  }
                    180 { $gwd=$page.Width.Point;  $ght=$page.Height.Point; }
                    270 { $gwd=$page.Height.Point; $ght=$page.Width.Point;  }
                }
                # PDF描画
                $wd = [System.Math]::Min($gwd, $ght) * ($dStumpPosSZ / 100.0)
                $ht = [System.Math]::Min($gwd, $ght) * ($dStumpPosSZ / 100.0)
                $px = $gwd * ($dStumpPosX / 100.0) - ($wd / 2.0)
                $py = $ght * ($dStumpPosY / 100.0) - ($ht / 2.0)
                $rt = New-Object PdfSharp.Drawing.XRect($px, $py, $wd, $ht)
                $gfx.DrawImage($pdfstm, $rt)
            }
        }
        $pdfdoc.Save($sDstPath)
    }
    end {}
}

try {
    $null = Write-Host "---PDFDateStamp---"
    $sTmpPath = [System.IO.Path]::GetTempFileName()
    MakePDFDateStamp $sTmpPath $sTopTxt $sMdlTxt $sBtmTxt
    AnnotatePDFStamp $sSrcPath $sDstPath $lTgtPageStt $lTgtPageEnd $sTmpPath $dStumpPosX $dStumpPosY $dStumpPosSZ 
}
catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
} finally {
    Remove-Item -LiteralPath $sTmpPath
}
