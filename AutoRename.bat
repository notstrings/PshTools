@echo off
pushd %~dp0
chcp 65001
"C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -File ".\AutoRename.ps1" %*
popd

