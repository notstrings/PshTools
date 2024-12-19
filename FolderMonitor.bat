@echo off
pushd %~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File ".\FolderMonitor.ps1" %*
popd

pause