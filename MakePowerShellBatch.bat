@echo off
pushd %~dp0
powershell -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File ".\MakePowerShellBatch.ps1" %*
popd

