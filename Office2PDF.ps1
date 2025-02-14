$ErrorActionPreference = "Stop"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:Office2PDF([System.IO.FileInfo] $Target) {
    $dname = [System.IO.Path]::GetDirectoryName($Target.FullName)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($Target.FullName)
    $ename = [System.IO.Path]::GetExtension($Target.FullName)
    $srcpath = $Target.FullName
    $dstpath = [System.IO.Path]::Combine($dname, $fname + ".pdf")
    switch ($ename.ToLower()) {
        ({$_ -eq ".doc" -or $_ -eq ".docx" -or $_ -eq ".dotm"}) {
            # Word
            try {
                $AppDOC = New-Object -ComObject Word.Application
                $AppDOC.Visible = $true
                $AppDOC.DisplayAlerts = 0 # wdAlertsNone
                $null = $document = $AppDOC.Documents.Open($srcpath)
                $null = $document.ExportAsFixedFormat(
                    $dstpath,                   # 出力ファイル名
                    17,                         # wdExportFormatPDF
                    $false,                     # OpenAfterExport
                    0,                          # wdExportOptimizeForPrint
                    0,                          # wdExportAllDocument
                    0,                          # From
                    0,                          # To
                    0,                          # wdExportDocumentContent
                    $false,                     # IncludeDocProps
                    $false,                     # KeepIRM
                    2,                          # wdExportCreateWordBookmarks
                    $true,                      # DocStructureTags
                    $true,                      # BitmapMissingFonts
                    $false                      # UseISO19005_1
                )
            } finally {
                if($null -ne $document){
                    $null = $document.Close(0)
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document)
                }
                if($null -ne $AppDOC){
                    $null = $AppDOC.Quit()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppDOC)
                }
            }
        }
        ({$_ -eq ".xls" -or $_ -eq ".xlsx" -or $_ -eq ".xlsm"}) {
            # Excel
            try {
                $AppXLS = New-Object -ComObject Excel.Application
                $AppXLS.Visible = $true
                $AppXLS.DisplayAlerts = $false
                $null = $workbook = $AppXLS.Workbooks.Open($srcpath)
                $null = $workbook.PrintOut(
                    [System.Type]::Missing,     # 印刷開始ページ番号
                    [System.Type]::Missing,     # 印刷終了ページ番号
                    [System.Type]::Missing,     # 印刷部数
                    [System.Type]::Missing,     # 印刷プレビューの有無
                    "Microsoft Print to PDF",   # 印刷プリンター
                    $true,                      # PrintToFile
                    [System.Type]::Missing,     # 部単位で印刷
                    $dstpath                    # 印刷するファイルの名前
                )
            } finally {
                if($null -ne $workbook){
                    $null = $workbook.Close($false)
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
                }
                if($null -ne $AppXLS){
                    $null = $AppXLS.Quit()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppXLS)
                }
            }
        }
        ({$_ -eq ".ppt" -or $_ -eq ".pptx" -or $_ -eq ".pptm"}) {
            # PowerPoint
            # ・PowerPoint2010の場合やたらと不安定だが...まぁ無理もねぇかと放置
            try {
                $AppPPT = New-Object -ComObject PowerPoint.Application
                $AppPPT.Visible = -1 # msoTrue=-1
                $AppPPT.DisplayAlerts = $false
                $null = $presentation = $AppPPT.Presentations.Open($srcpath)
                $null = $presentation.SaveAs(
                    $dstpath,                   # FileName
                    32,                         # ppSaveAsPDF=32
                    -1                          # EmbedTrueTypeFonts
                )
            } finally {
                if($null -ne $presentation){
                    $null = $presentation.Close()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
                }
                if($null -ne $AppPPT){
                    $null = $AppPPT.Quit()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppPPT)
                }
            }
        }
        ({$_ -eq ".vsd" -or $_ -eq ".vsdx" -or $_ -eq ".vsdm"}) {
            # Visio
            try {
                $AppVSD = New-Object -ComObject Visio.Application
                $AppVSD.Visible = $true
                $null = $document = $AppVSD.Documents.Open($srcpath)
                $null = $document.ExportAsFixedFormat(
                    1,                          # visFixedFormatPDF
                    $dstpath,                   # 出力ファイル名
                    1,                          # 出力品質
                    0                           # ページ範囲
                )
            } finally {
                if($null -ne $document){
                    $null = $document.Close()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document)
                }
                if($null -ne $AppVSD){
                    $null = $AppVSD.Quit()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppVSD)
                }
            }
        }
        default {
            Write-Host "未対応のファイル形式: $srcpath" -ForegroundColor Yellow
        }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---$Title---"
	# 引数確認
    if ($args.Count -eq 0) {
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            Office2PDF (Get-Item $arg)
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
