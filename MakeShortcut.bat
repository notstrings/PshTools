@echo off
chcp 65001 > nul
pushd %~dp0
"C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -WindowStyle hidden -ExecutionPolicy RemoteSigned -File ".\MakeShortcut.ps1" %*
popd

