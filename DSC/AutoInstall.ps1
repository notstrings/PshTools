
# 管理者権限セッションでの実行を要求する
function isInAdmin {
    $UID = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $UID.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (isInAdmin -eq $false) {
    Start-Process powershell -ArgumentList "-NoExit", "-File", """$($MyInvocation.MyCommand.Path)""" -Verb RunAs
    exit 0
}

# Scoop
if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
}
scoop bucket add extras
scoop install 7zip
scoop install pandoc
scoop install doxygen
scoop install imagemagick
scoop install ghostscript
scoop install plantuml
scoop install nuget
scoop install gh

# Winget
winget configure --disable-interactivity "$($PSScriptRoot)\General.dsc.yaml"
winget configure --disable-interactivity "$($PSScriptRoot)\Develop.dsc.yaml"
