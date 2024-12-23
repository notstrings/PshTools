$ErrorActionPreference = "Stop"

# ファイル
function local:ExecFile([System.IO.FileInfo] $Target) {
    ReduceFile $Target.FullName
}

# フォルダ
function local:ExecDir([System.IO.DirectoryInfo] $Target) {
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ExecFile $elm
    }
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ExecDir  $elm
    }
    ReduceDir $Target.FullName
}

# 整理
function local:ReduceDir([string]$Target) {
    # 空フォルダ
    if ( @(Get-ChildItem -LiteralPath $Target -File     ).Length -eq 0 -and
         @(Get-ChildItem -LiteralPath $Target -Directory).Length -eq 0 ) {
        Remove-Item -LiteralPath $Target -Recurse -Force
        return
    }
    # フォルダ名
    # $remname = @(
    #     "XXX"
    # )
    # if ($remname -contains [System.IO.Path]::GetFileName($Target)) {
    #     Remove-Item -LiteralPath $Target -Recurse -Force
    #     return
    # }
    # 引き上げ
    # if ( @(Get-ChildItem -LiteralPath ([System.IO.Path]::GetDirectoryName($Target)) -File     ).Length -eq 0 -and
    #      @(Get-ChildItem -LiteralPath ([System.IO.Path]::GetDirectoryName($Target)) -Directory).Length -eq 1 ) {
    #     if((Test-Path -LiteralPath  [System.IO.Path]::Combine($Target, [System.IO.Path]::GetFileName($Target)) -eq $false)){
    #         Move-Item -Path ($Target+"/*") ($Target+"/..") -Force
    #         Remove-Item -LiteralPath $Target -Recurse -Force
    #     }
    #     return
    # }
}
function local:ReduceFile([string]$Target) {
    # ファイル名
    $remname = @(
        "Thumbs.db", ".DS_Store"
    )
    if ($remname -contains [System.IO.Path]::GetFileName($Target)) {
        Remove-Item -LiteralPath $Target -Recurse -Force
        return
    }
    # 拡張子
    # $remexts = @(
    #     ".exe", ".dll", ".bat", ".txt", ".js", ".vbs", 
    #     ".url", ".lnk", ".html", ".htm", ".css", ".mht", 
    #     ".wav", ".mp3", ".ogg"
    # )
    # if ($remexts -contains [System.IO.Path]::GetExtension($Target)) {
    #     Remove-Item -LiteralPath $Target -Recurse -Force
    #     return
    # }
}

if ($args.Length -eq 0) {
    exit
}
$null = Write-Host "---ReduceDir---"
ForEach ($arg in $args) {
    if( Test-Path -LiteralPath $arg ){
        if ((Get-Item $arg).PSIsContainer) {
            ExecDir (Get-Item $arg)
        }
    }
}
$null = Write-Host "<End>"
