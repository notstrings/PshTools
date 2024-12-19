. "$($PSScriptRoot)/ModuleMisc.ps1"

Add-Type -Name ConsoleAPI -Namespace Win32Util -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
'

$hwnd = [Win32Util.ConsoleAPI]::GetConsoleWindow()
Write-Host $hwnd