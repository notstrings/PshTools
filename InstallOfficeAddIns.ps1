$ErrorActionPreference = "Stop"

$Title = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

function InstallOfficeAddIns([System.IO.FileInfo] $Target) {
    $dname = [System.IO.Path]::GetDirectoryName($Target.FullName)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($Target.FullName)
    $ename = [System.IO.Path]::GetExtension($Target.FullName).ToLower()
    switch ($ename) {
        # Word
        ".dot"  { AddWordAddIn $dname $fname $ename }
        ".dotm" { AddWordAddIn $dname $fname $ename }
        # Excel
        ".xla"  { AddExcelAddIn $dname $fname $ename }
        ".xlam" { AddExcelAddIn $dname $fname $ename }
        # PowerPoint
        ".ppa"  { AddPowerPointAddIn $dname $fname $ename }
        ".ppam" { AddPowerPointAddIn $dname $fname $ename }
        # Visio
        ".vsd"  { AddVisioAddIn $dname $fname $ename }
        ".vsdm" { AddVisioAddIn $dname $fname $ename }
        default {
            Write-Host "未対応のファイル形式: $($Target.FullName)" -ForegroundColor Yellow
        }
    }
}

function AddWordAddIn($dname, $fname, $ename) {
    $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
    $dstpath = "$($env:APPDATA)\Microsoft\Word\STARTUP\$($fname + $ename)"
    New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    Copy-Item $srcpath $dstpath -Force
}

function AddExcelAddIn($dname, $fname, $ename) {
    $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
    $dstpath = "$($env:APPDATA)\Microsoft\Excel\AddIns\$($fname + $ename)"
    New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    Copy-Item $srcpath $dstpath -Force
    # COM登録
    try {
        $AppXLS = New-Object -ComObject Excel.Application
        $XLSDOC = $AppXLS.Workbooks.Add()
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
        if($null -ne $XLSDOC){
            $null = $XLSDOC.Close($false)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($XLSDOC)
        }
        if($null -ne $AppXLS){
            $null = $AppXLS.Quit()
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppXLS)
        }
    }
}

function AddPowerPointAddIn($dname, $fname, $ename) {
    $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
    $dstpath = "$($env:APPDATA)\Microsoft\AddIns\$($fname + $ename)"
    New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    Copy-Item $srcpath $dstpath -Force
    # COM登録
    try {
        $AppPPT = New-Object -ComObject PowerPoint.Application
        $PPTDOC = $AppPPT.Presentations.Add()
        $PPTADIS = $AppPPT.AddIns
        $bFind = $PPTADIS | Where-Object { $_.Name -eq ($fname + $ename) }
        if (-not $bFind) {
            $PPTADIN = $PPTADIS.Add($dstpath, $false)
            $PPTADIN.Installed = $true
        }
    } finally {
        if($null -ne $PPTADIN){
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PPTADIN)
        }
        if($null -ne $PPTDOC){
            $null = $PPTDOC.Close($false)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PPTDOC)
        }
        if($null -ne $AppPPT){
            $null = $AppPPT.Quit()
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppPPT)
        }
    }
}

function AddVisioAddIn($dname, $fname, $ename) {
    $srcpath = [System.IO.Path]::Combine($dname, $fname + $ename)
    $dstpath = "$($env:APPDATA)\Microsoft\Visio\AddIns\$($fname + $ename)"
    New-Item ([System.IO.Path]::GetDirectoryName($dstpath)) -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    Copy-Item $srcpath $dstpath -Force
    # COM登録
    try {
        $AppVIS = New-Object -ComObject Visio.Application
        $VISDOC = $AppVIS.Documents.Add("")
        $VISADIS = $AppVIS.AddIns
        $bFind = $VISADIS | Where-Object { $_.Name -eq ($fname + $ename) }
        if (-not $bFind) {
            $VISADIN = $VISADIS.Add($dstpath, $false)
            $VISADIN.Installed = $true
        }
    } finally {
        if($null -ne $VISADIN){
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($VISADIN)
        }
        if($null -ne $VISDOC){
            $null = $VISDOC.Close()
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($VISDOC)
        }
        if($null -ne $AppVIS){
            $null = $AppVIS.Quit()
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($AppVIS)
        }
    }
}

# メイン処理
try {
    Write-Host "---$Title---"

    if ($args.Count -eq 0) { exit }

    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            InstallOfficeAddIns (Get-Item $arg)
        }
    }
} catch {
    Write-Host "---例外発生---"
    Write-Host $_.Exception.Message
    Write-Host $_.ScriptStackTrace
    Write-Host "--------------"
}
