@echo off
pushd %~dp0
chcp 65001
"C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -File ".\GiteaInit.ps1" %*
popd

