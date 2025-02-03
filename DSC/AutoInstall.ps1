
# 管理者権限セッションでの実行を要求する
function isInAdmin {
    $UID = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $UID.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (isInAdmin -eq $false) {
    Start-Process powershell -ArgumentList "-NoExit", "-File", """$($MyInvocation.MyCommand.Path)""" -Verb RunAs
    exit 0
}

## WinGet #####################################################################

# PSGalleryを信頼する
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# DSC Modulesをインストール
# Find-Module -Name NetworkingDSC  -Repository PSGallery | Install-Module
# Find-Module -Name 7ZipArchiveDsc -Repository PSGallery | Install-Module
# Find-Module -Name FileContentDsc -Repository PSGallery | Install-Module

# DSC実行
# winget configure --disable-interactivity "$($PSScriptRoot)\Environment.dsc.yaml"
winget configure --disable-interactivity "$($PSScriptRoot)\Packages.dsc.yaml"

# RoboCopy
# robocopy "Source" "Destination" /MIR /FFT /DCOPY:DAT /R:3 /W:5 /NFL /NP /XJ 
# robocopy "Source" "Destination" /MIR /FFT /DCOPY:DAT /R:3 /W:5 /NFL /NP /XJ 

## Scoop ######################################################################

if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
}
scoop bucket add extras
scoop install 7zip
scoop install pandoc
scoop install doxygen
scoop install imagemagick
scoop install ghostscript
scoop install qpdf
scoop install nuget
scoop install gh
