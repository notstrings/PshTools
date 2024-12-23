$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:MkPshBat([string] $TargetPath, [string] $Mode) {
    MakeBatch  $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("Shift-JIS"))
    ConvPshEnc $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeBatch([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ename = [System.IO.Path]::GetExtension($TargetPath)
    $tpath = [System.IO.Path]::GetFileName($TargetPath)
    $bpath = [System.IO.Path]::Combine($dname, $fname + ".bat")
    switch ($Mode) {
        "PSH5 CUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + "chcp 65001" + "`r`n"
                $text = $text + """$($ENV:SystemRoot)\system32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -ExecutionPolicy RemoteSigned -File "".\$tpath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH5 GUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + "chcp 65001" + "`r`n"
                $text = $text + """$($ENV:SystemRoot)\system32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -WindowStyle hidden -ExecutionPolicy RemoteSigned -File "".\$tpath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH5 ISE" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + "chcp 65001" + "`r`n"
                $text = $text + "start ""$($ENV:SystemRoot)\system32\WindowsPowerShell\v1.0\powershell_ise.exe"" -NoProfile -File "".\$tpath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "WINGET DSC" {
            if ($ename.ToLower() -eq ".yaml" -or $ename.ToLower() -eq ".dsc") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + "chcp 65001" + "`r`n"
                $text = $text + """$($ENV:USERPROFILE)\AppData\Local\Microsoft\WindowsApps\winget.exe"" configure "".\$tpath""`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
    }
    [IO.File]::WriteAllLines($bpath, $text, $Encoding)
}

function local:ConvPshEnc([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $text = [System.IO.File]::ReadAllLines($TargetPath, (AutoGuessEncodingSimple($TargetPath)))
    [System.IO.File]::WriteAllLines($TargetPath, $text, $Encoding)
}

try {
    $ret = ShowFileListDialogWithOption -Title "出力選択" -Message "対象ファイルをD&Dしてください" -FileList $args -Options @("PSH5 CUI", "PSH5 GUI", "PSH5 ISE", "WINGET DSC")
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
