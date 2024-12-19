$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:MkPshBat([string] $TargetPath, [string] $Mode) {
    MakeBatch  $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("Shift-JIS"))
    ConvPshEnc $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeBatch([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ppath = [System.IO.Path]::GetFileName($TargetPath)
    $bpath = [System.IO.Path]::Combine($dname, $fname + ".bat")
    switch ($Mode) {
        "CUI" {
            $text = ""
            $text = $text + "@echo off" + "`r`n"
            $text = $text + "pushd %~dp0" + "`r`n"
            $text = $text + "chcp 65001" + "`r`n"
            $text = $text + "$($ENV:systemroot)\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "".\$ppath"" %*" + "`r`n"
            $text = $text + "popd" + "`r`n"
        }
        "GUI" {
            $text = ""
            $text = $text + "@echo off" + "`r`n"
            $text = $text + "pushd %~dp0" + "`r`n"
            $text = $text + "chcp 65001" + "`r`n"
            $text = $text + "$($ENV:systemroot)\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File "".\$ppath"" %*" + "`r`n"
            $text = $text + "popd" + "`r`n"
        }
        "ISE" {
            $text = ""
            $text = $text + "@echo off" + "`r`n"
            $text = $text + "pushd %~dp0" + "`r`n"
            $text = $text + "chcp 65001" + "`r`n"
            $text = $text + "start $($ENV:systemroot)\system32\WindowsPowerShell\v1.0\powershell_ise.exe -NoProfile -File "".\$ppath"" %*" + "`r`n"
            $text = $text + "popd" + "`r`n"
        }
    }
    [IO.File]::WriteAllLines($bpath, $text, $Encoding)
}

function local:ConvPshEnc([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $text = [System.IO.File]::ReadAllLines($TargetPath, (AutoGuessEncodingSimple($TargetPath)))
    [System.IO.File]::WriteAllLines($TargetPath, $text, $Encoding)
}

try {
    $ret = ShowFileListDialogWithOption -Title "出力選択" -Message "対象PS1ファイルをD&Dしてください" -FileList $args -FIleFilter "\.ps1$" -Options @("CUI", "GUI", "ISE")
    if ($ret[0] -eq "OK") {
        foreach($elm in $ret[1]) {
            MkPshBat $elm $ret[2]
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
