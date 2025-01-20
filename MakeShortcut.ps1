$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

function local:MkLink([string] $TargetPath, [string] $Mode) {
    $dname = [System.IO.Path]::GetDirectoryName($TargetPath)
    $fname = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ename = [System.IO.Path]::GetExtension($TargetPath)
    $exepath = [System.IO.Path]::Combine($dname, $fname + $ename)
    switch ($Mode) {
        "Current" { $lnkpath = [System.IO.Path]::Combine($dname, $fname + ".lnk") }
        "SendTo"  { $lnkpath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\SendTo", $fname + ".lnk")  }
        "StartUp" { $lnkpath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup", $fname + ".lnk") }
    }
    switch ($ename.ToLower()) {
        ".exe" {
            $exepath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $WSH = New-Object -ComObject WScript.Shell
            $lnk = $WSH.CreateShortCut($lnkpath)
            $lnk.TargetPath       = "$exepath"
            $lnk.IconLocation     = "$exepath"
            $lnk.Arguments        = ""
            $lnk.WorkingDirectory = """$dname"""
            $null = $lnk.Save()
        }
        ".bat" {
            $exepath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $WSH = New-Object -ComObject WScript.Shell
            $lnk = $WSH.CreateShortCut($lnkpath)
            $lnk.TargetPath       = "cmd.exe"
            $lnk.IconLocation     = "$($ENV:SystemRoot)\System32\Imageres.dll, 262"
            $lnk.Arguments        = "/C ""$exepath"""
            $lnk.WorkingDirectory = """$dname"""
            $null = $lnk.Save()
        }
        ".ps1" {
            $exepath = [System.IO.Path]::Combine($dname, $fname + $ename)
            $WSH = New-Object -ComObject WScript.Shell
            $lnk = $WSH.CreateShortCut($lnkpath)
            $lnk.TargetPath       = """$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe"""
            $lnk.IconLocation     = "$($ENV:SystemRoot)\System32\WindowsPowerShell\v1.0\powershell.exe, 0"
            $lnk.Arguments        = "-ExecutionPolicy RemoteSigned ""$exepath"""
            $lnk.WorkingDirectory = """$dname"""
            $null = $lnk.Save()
        }
    }
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---$Title---"
    $ret = ShowFileListDialogWithOption `
            -Title $Title `
            -Message "対象ファイル(exe/bat/ps1)をドラッグ＆ドロップしてください" `
            -FileList $args `
            -FileFilter "\.(exe|bat|ps1)$" `
            -Options @("Current", "SendTo", "StartUp")
    if ($ret[0] -eq "OK") {
        foreach($elm in $ret[1]) {
            if (Test-Path -LiteralPath $elm) {
                MkLink $elm $ret[2]
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
