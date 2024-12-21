@echo off
pushd %~dp0
chcp 65001
"C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File ".\MakeBatch.ps1" %*
popd

