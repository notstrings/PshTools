# PshTools

自分で一個づつコマンド実行するのが面倒くさいので
よくやるヤツをBatにD&Dすればいーよーにした物体です

## 準備

```psh
Set-ExecutionPolicy RemoteSigned -scope CurrentUser
invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
scoop bucket add extras
scoop install imagemagick
scoop install ghostscript
scoop install qpdf
```

PDF日付印についてはPDFSharpとrsvg-convertを使うので``PDFDateStamp.ps1``のSetup関数を実行してください

## 説明

頑張れ人工知能
https://deepwiki.com/notstrings/PshTools
