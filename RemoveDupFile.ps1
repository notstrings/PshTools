$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:RemoveDupFile([string[]] $Targets) {
    $hash = @{}
    $Targets.GetEnumerator() |
    ForEach-Object {
        Get-ChildItem -LiteralPath $_ -File |
        ForEach-Object {
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
            if (-not $hash.ContainsKey($fname)){
                $hash[$fname] = @()
            }
            $hash[$fname] += $_
        }
    }
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
