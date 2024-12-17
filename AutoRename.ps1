$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

# ファイル名の処理
function local:CleanupFName([System.IO.FileInfo] $Target) {
    CleanupNodeName $Target.FullName $Target.LastWriteTime $false
}

# フォルダ名の処理
function local:CleanupDName([System.IO.DirectoryInfo] $Target) {
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        CleanupDName $elm
    }
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        CleanupFName $elm
    }
    CleanupNodeName $Target.FullName $Target.CreationTime $true
}

# ファイル・フォルダ名の処理
function local:CleanupNodeName([string] $TargetPath, [datetime] $TargetDate, [bool] $isDir) {
    try {
        # 修正前名称
        $srcpath = $TargetPath
        # 修正後名称
        $dstpath = $TargetPath
        if ($isDir -eq $false) {
            $dname = [System.IO.Path]::GetDirectoryName($dstpath)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($dstpath)
            $ename = [System.IO.Path]::GetExtension($dstpath)
        } else {
            $dname = [System.IO.Path]::GetDirectoryName($dstpath)
            $fname = [System.IO.Path]::GetFileName($dstpath)
            $ename = ""
        }
        $fname = RestrictTextZen    -Text $fname -Chars "Ａ-Ｚａ-ｚ０-９　（）［］｛｝"
        $fname = RestrictTextHan    -Text $fname
        $fname = RestrictTextDate   -Text $fname -Format "yyyyMMdd" -RefDate $TargetDate
        $fname = RestrictTextBlank  -Text $fname
        $ename = RestrictTextBlank  -Text $ename
        $dstpath = [System.IO.Path]::Combine($dname, $fname + $ename)
        # 必要があればリネーム
        if ($fname -ne "") {
            if ($srcpath -ne $dstpath) {
                $null = Write-Host "---"
                $null = Write-Host "src : $srcpath"
                $null = Write-Host "dst : $dstpath"
                $null = MoveItemWithUniqName $srcpath $dstpath
            }
        }
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

try {
    if ($args.Length -eq 0) {
        exit 1
    }
    $null = Write-Host "<<Start>>"
    ForEach ($arg in $args) {
        if( Test-Path -LiteralPath $arg ){
            if ((Get-Item $arg).PSIsContainer) {
                CleanupDName (Get-Item $arg)
            } else {
                CleanupFName (Get-Item $arg)
            }
        }
    }
    $null = Write-Host "<<End>>"
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
