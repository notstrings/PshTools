$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:RemoveDupFile([string[]] $Targets) {
    $hash = @{}
    $Targets |
    ForEach-Object {
        Get-ChildItem -LiteralPath $_ -File |
        ForEach-Object {
            $uniqkey = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
            if (-not $hash.ContainsKey($uniqkey)){
                $hash[$uniqkey] = @()
            }
            $hash[$uniqkey] += $_
        }
    }
    # ForEach-Object {
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

if ($args.Length -eq 0) {
    exit
}
$null = Write-Host "---ReduceDir---"
RemoveDupFile $args
