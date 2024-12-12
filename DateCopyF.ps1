$ErrorActionPreference = "Stop"

# ファイルコピー
function local:CopyItem([string] $SrcName, [string] $DstName, [bool] $isDir) {
    # ユニーク名取得
    $sUniq = $DstName
    $lUniq = 1
    while( (Test-Path -LiteralPath $sUniq) ) {
        if ($isDir -eq $false) {
            $dname = [System.IO.Path]::GetDirectoryName($DstName)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($DstName)
            $ename = [System.IO.Path]::GetExtension($DstName)
        }else{
            $dname = [System.IO.Path]::GetDirectoryName($DstName)
            $fname = [System.IO.Path]::GetFileName($DstName)
            $ename = ""
        }
        $sUniq = [System.IO.Path]::Combine($dname, $fname + " ($lUniq)" + $ename)
        $lUniq++
    }
    # 進捗付きコピー
    if ($isDir -eq $false) {
        $index = 0
        $count = 1
    } else {
        $index = 0
        $count = (Get-ChildItem $SrcName -Recurse).Length
    }
    Copy-Item -LiteralPath $SrcName -Destination $sUniq -PassThru -Recurse | 
    ForEach-Object {
        Write-Progress "$fname" -PercentComplete (($index / $count)*100)
        if ($index -lt $count){
            $index += 1
        }
    } | Out-Null
    Write-Host "$count files copied."
}

function local:DateCopyFile([System.IO.FileInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($spath)
    $ename = [System.IO.Path]::GetExtension($spath)
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    CopyItem $spath $dpath $false
}

function local:DateCopyDir([System.IO.DirectoryInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileName($spath)
    $ename = ""
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    CopyItem $spath $dpath $true
}

try {
    if ($args.Length -eq 0) {
        exit
    }
    $null = Write-Host "---DateCopy---"
    ForEach ($arg in $args) {
        if( Test-Path -LiteralPath $arg ){
            if ((Get-Item $arg).PSIsContainer) {
                DateCopyDir  (Get-Item $arg)
            } else {
                DateCopyFile (Get-Item $arg)
            }
        }
    }
    cmd /c timeout 3
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
