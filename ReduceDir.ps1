$ErrorActionPreference = "Stop"

# ファイル
function local:ReduceFile([System.IO.FileInfo] $Target) {
    ReduceFile $Target.FullName
}

# フォルダ
function local:ReduceDir([System.IO.DirectoryInfo] $Target) {
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ReduceFile $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ReduceDir  $elm
    }
    Reduce $Target.FullName
}

# 整理
function local:Reduce([string]$Target) {
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

try {
    $null = Write-Host "---ReduceDir---"
    if ($args.Length -eq 0) {
        exit
    }
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                ReduceDir (Get-Item $arg)
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
