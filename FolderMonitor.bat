@echo off
pushd %~dp0
powershell -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File ".\FolderMonitor.ps1" %*
popd

