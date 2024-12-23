@echo off
pushd %~dp0
chcp 65001
"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -File ".\test2.ps1" %*
popd

