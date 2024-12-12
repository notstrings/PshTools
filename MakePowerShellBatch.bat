@echo off
pushd %~dp0
powershell  -WindowStyle hidden -ExecutionPolicy Bypass -File ".\MakePowerShellBatch.ps1" %*
popd

