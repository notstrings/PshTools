﻿$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

# ファイル
function local:AutoRenameFile([System.IO.FileInfo] $Target) {
    AutoRename $Target.FullName $Target.LastWriteTime $false
}

# フォルダ
function local:AutoRenameDir([System.IO.DirectoryInfo] $Target) {
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        AutoRenameDir  $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        AutoRenameFile $elm
    }
    AutoRename $Target.FullName $Target.CreationTime $true
}

# ファイル・フォルダ名の処理
function local:AutoRename([string] $TargetPath, [datetime] $TargetDate, [bool] $isDir) {
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

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---AutoRename---"
	# 引数確認
    if ($args.Length -eq 0) {
        exit 1
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                AutoRenameDir  (Get-Item $arg)
            } else {
                AutoRenameFile (Get-Item $arg)
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
