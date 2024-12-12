@echo off
pushd %~dp0
powershell  -WindowStyle hidden -ExecutionPolicy Bypass -File ".\MakePowerShellLink.ps1" %*
popd

