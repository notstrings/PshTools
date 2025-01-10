$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:DateCopyFile([System.IO.FileInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($spath)
    $ename = [System.IO.Path]::GetExtension($spath)
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $cdate + "_" + $fname + $ename)
    CopyItemWithUniqName $spath $dpath
}

function local:DateCopyDir([System.IO.DirectoryInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileName($spath)
    $ename = ""
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $cdate + "_" + $fname + $ename)
    CopyItemWithUniqName $spath $dpath
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
