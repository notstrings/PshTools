# MakePshHandler

MS先生はPowerShellスクリプトをD&Dで起動したいとゆーたったそれだけのためにどんだけ頑張らせるんだろーか<br>
ということでソレ用のバッチ生成とショートカットファイル生成を用意して無駄にリストボックスによるD&Dにも対応してみた<br>

## MakePowerShellBatch

``MakePowerShellBatch.bat``はカレントの``MakePowerShellBatch.ps1``を起こすだけです<br>
``MakePowerShellBatch.ps1``はフォーム起動してD&Dで指定した``.ps1``をUTF-8(BOM付)に変換して起動用バッチを吐きます<BR>

中身は以下の通りですんでBypassが嫌ならスクリプトを弄ってください<BR>
あとウィンドウを非表示にしたい場合は``-WindowStyle hidden``辺り追加してください<BR>
※一瞬見えるんだけどまぁいいでしょ<BR>

```
@echo off
pushd %~dp0
powershell -ExecutionPolicy Bypass -File ".\[[指定したPowerShell]]" %*
popd
```

## MakePowerShellLink

``MakePowerShellLink.bat``はカレントの``MakePowerShellLink.ps1``を起こすだけです<br>
``MakePowerShellLink.ps1``はフォーム起動してD&Dで指定した``.ps1``の起動用ショートカットファイルを生成します<br>

## メモ

昔々Shift-JISなバッチからUTF-8なpowershell起動すると文字化けしてた気がするんで<br>
必死こいて文字コード判別してやろーと思ったんだがいつの間にかそんなコト無くなってて泣いてる<br>
