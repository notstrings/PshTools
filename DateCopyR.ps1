﻿$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:DateCopyFile([System.IO.FileInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($spath)
    $ename = [System.IO.Path]::GetExtension($spath)
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    CopyItemWithUniqName $spath $dpath
}

function local:DateCopyDir([System.IO.DirectoryInfo] $Target) {
    $spath = $Target.FullName
    $dname = [System.IO.Path]::GetDirectoryName($spath)
    $fname = [System.IO.Path]::GetFileName($spath)
    $ename = ""
    $cdate = (Get-Date).ToString('yyyyMMdd')
    $dpath = [System.IO.Path]::Combine($dname, $fname + "_" + $cdate + $ename)
    CopyItemWithUniqName $spath $dpath
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
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
