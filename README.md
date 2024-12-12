# PshTools

バッチとかで作ってた小物ツールをいい加減そろそろPowerShellに移植するぜという事で作った代物<br>
2024現在何故かWin10でオシゴトしているのでWin10+PowerShell5で動作確認してます<br>

## ライセンス

とりあえずよく分かんないのでMITライセンスを名乗っておきます<br>
~~いやアレだろ、ミットってキャッチャーの持ってるあの～~~

## AutoRename

``AutoRename.ps1``に自動リネームを適用したいパスを食わせてください<br>
パスの指定は複数可能でディレクトリを指定した場合その配下の全てのファイルフォルダを処理します<br>
ファイル名が衝突した場合は``(1)``とか勝手に付けて衝突を回避します<BR>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

* ファイル・フォルダ名に含まれる全角英数と＋全角括弧/空白(``（）［］｛｝＿``)を半角にする
* ファイル・フォルダ名に含まれる複数の半角空白を一つにする
* ファイル・フォルダ名に含まれる半角カナを全角カナにする
* ファイル・フォルダ名に日付が含まれている場合、その部分を``YYYYMMDD``に直す
  * YYYY-MM-DD or YYYY.MM.DD
  * YYYY年MM月DD日
  * 和暦YY-MM-DD or 和暦YY.MM.DD 
  * 和暦YY年MM月DD日
  * YY-MM-DD or YY.MM.DD※1
  * YY年MM月DD日※1

※1:表記に西暦or和暦を張り付けてファイル変更日付orフォルダ作成日付に一致すればリネーム<br>

## DateCopy

``DateCopyX.ps1``に日付を適用したいパスを食わせてください※Fが前でRが後<br>
日付をファイル・フォルダ名の前or後に勝手に追加してコピーします<br>
ファイル名が衝突した場合は``(1)``とか勝手に付けて衝突を回避します<BR>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

バッチでセコセコやってた頃は``XCOPY``とか``ROBOCOPY``だったのにね<BR>

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

空白を含むパスに使うと駄目な場合は大人しく``MakePowerShellBatch``のほーをどーぞ<br>
