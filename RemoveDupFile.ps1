$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:RemoveDupFile([string[]] $Targets) {
    $hash = @{}
    $Targets | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -File |
        ForEach-Object {
            $uniqkey = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
            if (-not $hash.ContainsKey($uniqkey)){
                $hash[$uniqkey] = @()
            }
            $hash[$uniqkey] += $_
        }
    }
    # $Targets | ForEach-Object {
    #     Get-ChildItem -LiteralPath $_ -File |
    #     ForEach-Object {
    #         $uniqkey = Get-FileHash -LiteralPath $_.FullName -Algorithm MD5
    #         if (-not $hash.ContainsKey($uniqkey)){
    #             $hash[$uniqkey] = @()
    #         }
    #         $hash[$uniqkey] += $_
    #     }
    # }
    $hash.Values | ForEach-Object {
        $_ |
        Sort-Object -Property Length -Descending |
        Select-Object -Skip 1 | ForEach-Object { 
            MoveTrush -Path $_.FullName
        }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb")

try {
    $null = Write-Host "---$Title---"
	# 引数確認
    if ($args.Count -eq 0) {
        exit
    }
	# 処理実行
    RemoveDupFile $args
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
