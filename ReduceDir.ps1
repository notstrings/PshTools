$ErrorActionPreference = "Stop"

# ファイル
function local:ReduceFile([System.IO.FileInfo] $Target) {
    Reduce $Target.FullName $false
}

# フォルダ
function local:ReduceDir([System.IO.DirectoryInfo] $Target) {
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File)) {
        ReduceFile $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        ReduceDir  $elm
    }
    Reduce $Target.FullName $true
}

# 整理
function local:Reduce([string]$Target, [bool]$isDir) {
    if ($isDir) {
        # フォルダ
        ## 空フォルダ
        if ( @(Get-ChildItem -LiteralPath $Target -File     ).Length -eq 0 -and
            @(Get-ChildItem -LiteralPath $Target -Directory).Length -eq 0 ) {
            Remove-Item -LiteralPath $Target -Recurse -Force
            return
        }
        ## 不要フォルダ名
        # $remname = @(
        #     "XXX"
        # )
        # if ($remname -contains [System.IO.Path]::GetFileName($Target)) {
        #     Remove-Item -LiteralPath $Target -Recurse -Force
        #     return
        # }
        ## フォルダ引き上げ
        # if ( @(Get-ChildItem -Path ([System.IO.Path]::Combine($Target, "..")) -File     ).Length -eq 0 -and
        #      @(Get-ChildItem -Path ([System.IO.Path]::Combine($Target, "..")) -Directory).Length -eq 1 ) {
        #     $dup = ([System.IO.Path]::Combine($Target, [System.IO.Path]::GetFileName($Target)))
        #     if (Test-Path -LiteralPath $dup) {
        #         Move-Item -LiteralPath $dup (GenUniqName ($dup+"_") $true) -Force
        #     }
        #     Move-Item   -Path ($Target+"/*") ($Target+"/..") -Force
        #     Remove-Item -LiteralPath $Target -Recurse -Force
        #     return
        # }
    } else {
        # ファイル名
        ## 不要ファイル名
        $remname = @(
            "Thumbs.db", ".DS_Store"
        )
        if ($remname -contains [System.IO.Path]::GetFileName($Target)) {
            Remove-Item -LiteralPath $Target -Recurse -Force
            return
        }
        ## 不要拡張子
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
}

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---ReduceDir---"
	# 引数確認
    if ($args.Length -eq 0) {
        exit
    }
	# 処理実行
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
}
