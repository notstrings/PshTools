$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:MkPshLnk([string] $TargetPath) {
    MakeLink   $TargetPath
    ConvPshEnc $TargetPath ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeLink([string] $TargetPath) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ppath = [System.IO.Path]::Combine($dname, $fname + ".ps1")
    $lpath = [System.IO.Path]::Combine($dname, $fname + ".lnk")
    $WSH = New-Object -ComObject WScript.Shell
    $lnk = $WSH.CreateShortCut($lpath)
    $lnk.TargetPath       = """C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe"""
    $lnk.IconLocation     = "C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe, 0"
    $lnk.Arguments        = "-ExecutionPolicy RemoteSigned ""$ppath"""
    $lnk.WorkingDirectory = """$dname"""
    $null = $lnk.Save()
}

function local:ConvPshEnc([string] $TargetPath, [System.Text.Encoding] $Encoding) {
    $text = [System.IO.File]::ReadAllLines($TargetPath, (AutoGuessEncodingSimple($TargetPath)))
    [System.IO.File]::WriteAllLines($TargetPath, $text, $Encoding)
}

try {
    $ret = ShowFileListDialog -Title "出力選択" -Message "対象PS1ファイルをD&Dしてください" -FileList $args -FIleFilter "\.ps1$"
    if ($ret[0] -eq "OK") {
        foreach($elm in $ret[1]) {
            MkPshLnk $elm
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
 
