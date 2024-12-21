@echo off
pushd %~dp0
chcp 65001
"C:\Users\notst\AppData\Local\Microsoft\WindowsApps\winget.exe" configure ".\General.dsc.yaml"
popd

