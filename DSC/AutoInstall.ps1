
# 管理者権限セッションでの実行を要求する
function isInAdmin {
    $UID = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    return $UID.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (isInAdmin) {
    winget configure "$($PSScriptRoot)\General.dsc"
    winget configure "$($PSScriptRoot)\Develop.dsc"
} else {
    Start-Process powershell -ArgumentList "-NoExit", "-File", """$($MyInvocation.MyCommand.Path)""" -Verb RunAs
}
