@echo off
pushd %~dp0
powershell -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File ".\MakePowerShellLink.ps1" %*
popd
