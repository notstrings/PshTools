@echo off
pushd %~dp0
chcp 65001
start C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell_ise.exe -NoProfile -File ".\FolderMonitor.ps1" %*
popd

