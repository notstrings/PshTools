$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

function local:MkPshBat([string] $TargetPath) {
    MakeBatch  $TargetPath ([System.Text.Encoding]::GetEncoding("Shift-JIS"))
    ConvPshEnc $TargetPath ([System.Text.Encoding]::GetEncoding("UTF-8"))
}

function local:MakeBatch([string] $TargetPath, [System.Text.Encoding] $Encoding) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ppath = [System.IO.Path]::GetFileName($TargetPath)
    $bpath = [System.IO.Path]::Combine($dname, $fname + ".bat")
    $text = ""
    $text = $text + "@echo off" + "`n"
    $text = $text + "pushd %~dp0" + "`n"
    $text = $text + "powershell -ExecutionPolicy Bypass -File "".\$ppath"" %*" + "`n"
    $text = $text + "popd" + "`n"
    [IO.File]::WriteAllLines($bpath, $text, $Encoding)
}

function local:ConvPshEnc([string] $TargetPath, [System.Text.Encoding] $Encoding) {
    $text = [System.IO.File]::ReadAllLines($TargetPath, (AutoGuessEncodingSimple($TargetPath)))
    [System.IO.File]::WriteAllLines($TargetPath, $text, $Encoding)
}

try {
    $ret = ShowDDDialog -Title "出力選択" -Message "対象PS1ファイルをD&Dしてください" -List $args
    if ( $ret[0] -eq "OK" ) {
        foreach($elm in $ret[1]) {
            MkPshBat $elm
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
    pause
}
