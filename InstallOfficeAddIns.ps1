$ErrorActionPreference = "Stop"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:InstallOfficeAddIns([System.IO.FileInfo] $Target) {
    $dname = [System.IO.Path]::GetDirectoryName($Target.FullName)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($Target.FullName)
    $ename = [System.IO.Path]::GetExtension($Target.FullName)
    switch ($ename.ToLower()) {
        # Word
        ({$_ -eq ".dot" -or $_ -eq ".dotm"}) {
            $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $dstpath = "$($env:APPDATA)\Microsoft\Word\STARTUP\$($fname + $ename)"
            $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
            Write-Output "F" | xcopy /Y /F /K $srcpath $dstpath
        }
        # Excel
        ({$_ -eq ".xla" -or $_ -eq ".xlam"}) {
            try {
                $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
                $dstpath = "$($env:APPDATA)\Microsoft\Excel\AddIns\$($fname + $ename)"
                $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
                Write-Output "F" | xcopy /Y /F /K $srcpath $dstpath
                $AppXLS = New-Object -ComObject Excel.Application
                $XLSWKBK = $AppXLS.Workbooks.Add()
                $XLSADIS = $AppXLS.Addins
                $bFind = $XLSADIS | Where-Object { $_.Name -eq ($fname + $ename) }
                if (-not $bFind){
                    $XLSADIN = $XLSADIS.Add($dstpath, $false)
                    $XLSADIN.Installed = $true
                }
            } finally {
                if($null -ne $XLSADIN){
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLSADIN)
                }
                if($null -ne $XLSWKBK){
                    $null = $XLSWKBK.Close($false)
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLSWKBK)
                }
                if($null -ne $AppXLS){
                    $null = $AppXLS.Quit()
                    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppXLS)
                }
            }
        }
        # PowerPoint
        ({$_ -eq ".ppa" -or $_ -eq ".ppam"}) {
            # PowerPointマクロは良く分からんのでとりあえずそれっぽい場所にコピーするだけ
            $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $dstpath = "$($env:APPDATA)\Microsoft\AddIns\$($fname + $ename)"
            $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
            Write-Output "F" | xcopy /Y /F /K $srcpath $dstpath
        }
        # Visio
        ({$_ -eq ".vsd" -or $_ -eq ".vsdm"}) {
            # Visioマクロは良く分からんのでとりあえずそれっぽい場所にコピーするだけ
            $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $dstpath = "$($env:APPDATA)\Microsoft\Visio\AddIns\$($fname + $ename)"
            $null = New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue
            Write-Output "F" | xcopy /Y /F /K $srcpath $dstpath
        }
        default {
            Write-Host "未対応のファイル形式: $($Target.FullName)" -ForegroundColor Yellow
        }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\Utility.xlam")

try {
    $null = Write-Host "---$Title---"
	# 引数確認
    if ($args.Count -eq 0) {
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            InstallOfficeAddIns (Get-Item $arg)
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
