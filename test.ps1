. "$($PSScriptRoot)/ModuleMisc.ps1"

class AppSettings {
    [System.ComponentModel.Description("名前")]
    [string]$AppName
    [int]$Version
    [bool]$AutoUpdate
    [string]$LogFilePath
    [System.Diagnostics.SourceLevels]$LogLevel
}
$settings = [AppSettings]@{
    AppName = "My Application"
    Version = 1
    AutoUpdate = $true
    LogFilePath = "C:\app.log"
    LogLevel = [System.Diagnostics.SourceLevels]::Information
}
# $json = Get-Content -Path $sPath | ConvertFrom-Json
# $conf = [System.Convert]::ChangeType($psConf, ([Config]))
$ret = ShowSettingDialog "Title" $settings
if ($ret -eq "OK") {
    $edit
}