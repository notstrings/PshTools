# PshTools

自分でチマチマとコマンド実行するのが面倒くさいのでよくやるヤツをBatにD&Dすればいーよーにした物体です。

ScoopとかでインストールしたCUIコマンドでゴチャゴチャやりますんで使用には事前準備が必要...

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
