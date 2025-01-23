$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:MkPshBat([string] $TargetPath, [string] $Mode) {
    MakeBatch  $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("Shift-JIS"))
    ConvPshEnc $TargetPath $Mode ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeBatch([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ename = [System.IO.Path]::GetExtension($TargetPath)
    $exepath = [System.IO.Path]::Combine(".",    $fname + $ename)
    $batpath = [System.IO.Path]::Combine($dname, $fname + ".bat")
    switch ($Mode) {
        "PSH5 CUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "chcp 65001 > nul" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + """$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -ExecutionPolicy RemoteSigned -File ""$exepath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH7 CUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "chcp 65001 > nul" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + """$($env:ProgramFiles)\PowerShell\7\pwsh.exe"" -NoProfile -ExecutionPolicy RemoteSigned -File ""$exepath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH5 GUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "chcp 65001 > nul" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + """$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -WindowStyle hidden -ExecutionPolicy RemoteSigned -File ""$exepath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH7 GUI" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "chcp 65001 > nul" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + """$($env:ProgramFiles)\PowerShell\7\pwsh.exe"" -NoProfile -WindowStyle hidden -ExecutionPolicy RemoteSigned -File ""$exepath"" %*" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
        "PSH5 ISE" {
            if ($ename.ToLower() -eq ".ps1") {
                $text = ""
                $text = $text + "@echo off" + "`r`n"
                $text = $text + "chcp 65001 > nul" + "`r`n"
                $text = $text + "pushd %~dp0" + "`r`n"
                $text = $text + """$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell_ise.exe"" -File ""$exepath"" -NoProfile" + "`r`n"
                $text = $text + "popd" + "`r`n"
            }
        }
    }
    [IO.File]::WriteAllLines($batpath, $text, $Encoding)
}

function local:ConvPshEnc([string] $TargetPath, [string] $Mode, [System.Text.Encoding] $Encoding) {
    $enc = AutoGuessEncodingFileSimple $TargetPath
    $txt = [System.IO.File]::ReadAllLines($TargetPath, $enc)
    [System.IO.File]::WriteAllLines($TargetPath, $txt, $Encoding)
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---$Title---"
    $ret = ShowFileListDialogWithOption `
            -Title $Title `
            -Message "対象ファイルをドラッグ＆ドロップしてください" `
            -FileList $args `
            -FileFilter "\.(ps1)$" `
            -Options @("PSH5 CUI", "PSH7 CUI", "PSH5 GUI", "PSH7 GUI", "PSH5 ISE")
    if ($ret[0] -eq "OK") {
        foreach($elm in $ret[1]) {
            if (Test-Path -LiteralPath $elm) {
                MkPshBat $elm $ret[2]
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
