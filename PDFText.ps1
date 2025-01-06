param (
    [Parameter(Mandatory = $true)]  [string] $sSrcPath,
    [Parameter(Mandatory = $true)]  [string] $sDstPath,
    [Parameter(Mandatory = $false)] [long]   $lTgtPageStt = 1,
    [Parameter(Mandatory = $false)] [long]   $lTgtPageEnd = 0,
    [Parameter(Mandatory = $true)]  [string] $sAnnText,
    [Parameter(Mandatory = $false)] [double] $dAnnTextSize = 12,
    [Parameter(Mandatory = $false)] [string] $sAnnTextColor = "black",
    [Parameter(Mandatory = $false)] [double] $dAnnTextPosX = 0,
    [Parameter(Mandatory = $false)] [double] $dAnnTextPosY = 0
)

# セットアップ
function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install nuget

    # セットアップ先
    $ToolDir = [System.IO.Path]::Combine($PSScriptRoot, "tool")
    $null = New-Item $ToolDir -ItemType Directory -ErrorAction SilentlyContinue

    # PDFSharp
    $ToolPath01 = [System.IO.Path]::Combine($ToolDir, "PDFsharp.1.50.5147")
    if ((Test-Path $ToolPath01) -eq $false) {
        nuget install PdfSharp -Version 1.50.5147 -OutputDirectory $ToolDir -Source 'https://api.nuget.org/v3/index.json'
    }
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
.PARAMETER sAnnText
    描画文字列
.PARAMETER dAnnTextSize
    描画文字サイズ
.PARAMETER sAnnTextColor
    描画文字色
.PARAMETER dAnnTextPosX
    描画文字横軸の割合を指定します
    0%は左端で100%は右端に位置します
.PARAMETER dAnnTextPosY
    描画文字横軸の割合を指定します
    ページの短編を基準とした割合を指定します
.EXAMPLE
    AnnotatePDFText "$($ENV:USERPROFILE)\Desktop\src.pdf" "$($ENV:USERPROFILE)\Desktop\dst.pdf" 1 0 "text" 32 "black" 50 50
#>
function local:AnnotatePDFText {
    param (
        [Parameter(Mandatory = $true)]  [string] $sSrcPath,
        [Parameter(Mandatory = $true)]  [string] $sDstPath,
        [Parameter(Mandatory = $false)] [long]   $lTgtPageStt = 1,
        [Parameter(Mandatory = $false)] [long]   $lTgtPageEnd = 0,
        [Parameter(Mandatory = $true)]  [string] $sAnnText,
        [Parameter(Mandatory = $false)] [double] $dAnnTextSize = 12,
        [Parameter(Mandatory = $false)] [string] $sAnnTextColor = "black",
        [Parameter(Mandatory = $false)] [double] $dAnnTextPosX = 0,
        [Parameter(Mandatory = $false)] [double] $dAnnTextPosY = 0
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
        $pdfdoc = New-Object PdfSharp.Pdf.PdfDocument
        $pdffnt = [PdfSharp.Drawing.XFont]::new("Yu Mincho", $dAnnTextSize)
        for($idx = 1; $idx -le $pagecount; $idx++) {
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
                # 文字描画
                $tx = $sAnnText
                $tx = $tx.Replace("[cpage]", $idx)
                $tx = $tx.Replace("[epage]", $pagecount)
                ## 領域
                $px = $gwd * ($dAnnTextPosX / 100.0)
                $py = $ght * ($dAnnTextPosY / 100.0)
                $rt = New-Object PdfSharp.Drawing.XRect($px, $py, 0, 0)
                ## 色
                $bc = [System.Drawing.ColorTranslator]::FromHtml($sAnnTextColor)
                $br = New-Object PdfSharp.Drawing.XSolidBrush([PdfSharp.Drawing.XColor]::FromArgb($bc.A,$bc.R,$bc.G,$bc.B))
                ## 文字描画
                $gfx.DrawString($tx, $pdffnt, $br, $rt, [PdfSharp.Drawing.XStringFormats]::Center)
            }
        }
        $pdfdoc.Save($sDstPath)
    }
    end {}
}

try {
    $null = Write-Host "---PDFText---"
    AnnotatePDFText $sSrcPath $sDstPath $lTgtPageStt $lTgtPageEnd $sAnnText $dAnnTextSize $sAnnTextColor $dAnnTextPosX $dAnnTextPosY 
}
catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
