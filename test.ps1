. "$($PSScriptRoot)/ModuleMisc.ps1"

function BinaryFormatterClone {
    param (
        [Parameter(Mandatory=$true)] [Object] $src
    )
    $ms = New-Object System.IO.MemoryStream
    $binaryFormatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
    $binaryFormatter.Serialize($ms, $src)
    $ms.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
    return $binaryFormatter.Deserialize($ms)
}


class ConfChild {
    [System.ComponentModel.Description("監視名称")]
    [string]$MonName
    [System.ComponentModel.Description("監視位置")]
    [string]$MonPath
}
class Conf {
    [ConfChild[]]$ConfChild
}
$child1 = New-Object ConfChild -Property @{MonName = "監視名称1"; MonPath = "C:\Path\To\Mon1"}
$child2 = New-Object ConfChild -Property @{MonName = "監視名称2"; MonPath = "D:\Path\To\Mon2"}
$child3 = New-Object ConfChild -Property @{MonName = "監視名称3"; MonPath = "E:\Path\To\Mon3"}
$conf = New-Object Conf -Property @{ ConfChild = @($child1, $child2, $child3) }


# JSON にシリアライズ
$confJson = $conf | ConvertTo-Json -Depth 5
$confJson | Out-File -FilePath "config.json"

$confJson = Get-Content -Path "config.json" | Out-String
$confRestored = $confJson | ConvertFrom-Json
$confInstance = New-Object Conf 
$confInstance.ConfChild = $confRestored.ConfChild | ForEach-Object {
    New-Object ConfChild -Property @{MonName = $_.MonName; MonPath = $_.MonPath}
}

$confRestored.ConfChild | ForEach-Object {
    Write-Host "MonName: $($_.MonName), MonPath: $($_.MonPath)"
}

Write-Host ""