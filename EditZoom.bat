@echo off
chcp 65001 > nul
pushd %~dp0
"C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -File ".\EditZoom.ps1" %*
popd

