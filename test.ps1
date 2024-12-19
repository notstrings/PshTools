. "$($PSScriptRoot)/ModuleMisc.ps1"

Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($args -join " ", "title")
