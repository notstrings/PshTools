$ErrorActionPreference = "Stop"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:DiffWord([System.IO.FileInfo] $LHS, [System.IO.FileInfo] $RHS) {
    try {
        $AppDOC = New-Object -ComObject Word.Application
        $AppDOC.Visible = $true
        $AppDOC.DisplayAlerts = 0 # wdAlertsNone
        $DocLHS = $AppDOC.Documents.Open($LHS.FullName, $false, $true)
        $DocRHS = $AppDOC.Documents.Open($RHS.FullName, $false, $true)
        if ($LHS.LastWriteTime -le $RHS.LastWriteTime) {
            $DocLHS = $AppDOC.Documents.Open($LHS.FullName, $false, $true)
            $DocRHS = $AppDOC.Documents.Open($RHS.FullName, $false, $true)
        } else {
            $DocLHS = $AppDOC.Documents.Open($RHS.FullName, $false, $true)
            $DocRHS = $AppDOC.Documents.Open($LHS.FullName, $false, $true)
        }
        $null = $AppDOC.CompareDocuments(
            $DocLHS,        # OriginalDocument
            $DocRHS,        # RevisedDocument
            2,              # Destination wdCompareDestinationNew=2
            1,              # Granularity wdGranularityWordLevel=1
            $False,         # CompareFormatting
            $False,         # CompareCaseChanges
            $False,         # CompareWhitespace
            $True,          # CompareTables
            $False,         # CompareHeaders
            $False,         # CompareFootnotes
            $True,          # CompareTextboxes
            $True,          # CompareFields
            $False,         # CompareComments
            $False          # CompareMoves
        )
    } finally {
        if($null -ne $DocRHS){
            $null = $DocRHS.Close(0)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($DocRHS)
        }
        if($null -ne $DocLHS){
            $null = $DocLHS.Close(0)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($DocLHS)
        }
        # if($null -ne $AppDOC){ 
        #     $AppDOC.Quit(); $AppDOC = $null
        #     $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppDOC)
        # }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.docx", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb.docx")

try {
    $null = Write-Host "---$Title---"
	# 引数確認
    if ($args.Count -ne 2) {
        exit
    }
	# 処理実行
    $exist = $true
    $exist = $exist -and (Test-Path -LiteralPath $args[0])
    $exist = $exist -and (Test-Path -LiteralPath $args[1])
    if ($exist) {
        DiffWord (Get-Item $args[0]) (Get-Item $args[1])
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
