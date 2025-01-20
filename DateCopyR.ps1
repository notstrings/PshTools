$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 設定 #######################################################################

function local:DateCopyFile([System.IO.FileInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($spath)
    $ename = [System.IO.Path]::GetExtension($spath)
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    $uniq = GenUniqName $dpath $false
    Write-Output "F" | xcopy $spath $uniq /F /K
}

function local:DateCopyDir([System.IO.DirectoryInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileName($spath)
    $ename = ""
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    $uniq = GenUniqName $dpath $true
    robocopy $spath $uniq /MIR /FFT /DCOPY:DAT /R:3 /W:5 /NFL /NP /XJ 
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
            if ([System.IO.Directory]::Exists($arg)) {
                DateCopyDir  (Get-Item $arg)
            } else {
                DateCopyFile (Get-Item $arg)
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
