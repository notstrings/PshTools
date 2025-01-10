$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:MkPshLnk([string] $TargetPath) {
    MakeLink   $TargetPath ([System.Text.Encoding]::GetEncoding("UTF-8"))
    ConvPshEnc $TargetPath ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeLink([string] $TargetPath, [System.Text.Encoding] $Encoding) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ename = [System.IO.Path]::GetExtension($TargetPath)
    $abspath = [System.IO.Path]::Combine($dname, $fname + $ename)
    $dstpath = [System.IO.Path]::Combine($dname, $fname + ".lnk")
    $WSH = New-Object -ComObject WScript.Shell
    $lnk = $WSH.CreateShortCut($dstpath)
    $lnk.TargetPath       = """$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"""
    $lnk.IconLocation     = "$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe, 0"
    $lnk.Arguments        = "-ExecutionPolicy RemoteSigned ""$abspath"""
    $lnk.WorkingDirectory = """$dname"""
    $null = $lnk.Save()
}

function local:ConvPshEnc([string] $TargetPath, [System.Text.Encoding] $Encoding) {
    $enc = AutoGuessEncodingFileSimple $TargetPath
    $txt = [System.IO.File]::ReadAllLines($TargetPath, $enc)
    [System.IO.File]::WriteAllLines($TargetPath, $txt, $Encoding)
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---$Title---"
    $ret = ShowFileListDialog `
            -Title $Title `
            -Message "対象PS1ファイルをドラッグ＆ドロップしてください" `
            -FileList $args `
            -FileFilter "\.ps1$"
    if ($ret[0] -eq "OK") {
        foreach ($elm in $ret[1]) {
            if (Test-Path -LiteralPath $elm) {
                MkPshLnk $elm
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
