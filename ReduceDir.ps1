$ErrorActionPreference = "Stop"

# ファイル処理
function local:ReduceF([System.IO.FileInfo] $Target) {
    ReduceFile $Target.FullName
}

# フォルダ処理
function local:ReduceD([System.IO.DirectoryInfo] $Target) {
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ReduceF $elm
    }
    ForEach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ReduceD $elm
    }
    ReduceFolder $Target.FullName
}

# 整理
function local:ReduceFolder([string]$Target) {
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
    # if ( @(Get-ChildItem -LiteralPath ($Target+"/..") -File     ).Length -eq 0 -and
    #      @(Get-ChildItem -LiteralPath ($Target+"/..") -Directory).Length -eq 1 ) {
    #     Move-Item -Path ($Target+"/*") ($Target+"/..") -Force
    #     Remove-Item -LiteralPath $Target -Recurse -Force
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
            ReduceD (Get-Item $arg)
        }
    }
}
$null = Write-Host "<End>"
