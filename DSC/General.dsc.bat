@echo off
pushd %~dp0
chcp 65001
winget configure ".\General.dsc.yaml"
popd

