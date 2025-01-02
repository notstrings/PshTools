# PshTools

バッチとかで作ってた小物ツール一式＆それ用のPowerShellモジュール<br>
好き放題改造してその場で使ってるんでコンパイルが必要な言語を避けた結果のPowerShell<br>
汎用そうな機能は大体全部``ModuleMisc.ps1``に分割してますんで切った張ったはどーにでも<BR>

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

## Office2PDF

``Office2PDF.ps1``にPDFにしたいオフィス系ファイルのパスを食わせてください<br>
Word/Excel/PowerPoint/Visioで一通りの形式に対応しているような気がする<br>
漢らしくウィンドウ表示ありでPDF保存しにいきますんで処理完了まで待ってください<br>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

何かPowerPoint2010の場合挙動がおかしいんだが正直俺使わねぇからどーでもいいや（ぉぃ<br>

## DiffWord

``DiffWord.ps1``に比較したいWordファイルのパスを２つ食わせてください<br>
イチイチ考えるのが面倒くさい比較元・比較先を最終変更日時で判断して自動で比較処理を実施します<br>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

## ReduceDir

``ReduceDir.ps1``に整理したいフォルダのパスを食わせてください<br>
``Thumbs.DB``と``.DS_Store``削除/空フォルダ削除を実行します<br>
``ReduceFile``/``ReduceFOlder``関数改造すれば他にも色々できますねー<br>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

## RemoveDupFile

``RemoveDupFile.ps1``に重複ファイル削除したいフォルダのパスを複数食わせてください<br>
全フォルダのファイル名(拡張子無視)で重複チェックしてファイルサイズが一番デカイのを残します<br>
D&Dでパス指定したい場合は``MakePowerShell*.ps1``でバッチかショートカットファイルを作ってください<br>

## FolderMonitor

``FolderMonitor.ps1``は起動するとタスクトレイに常駐します<br>
右クリックメニューから設定して監視対象フォルダを指定してください<br>
デフォで５分に一回指定フォルダの差分を確認してIPMessengerで自分に通知します<br>
タスクトレイの右クリックメニューからGUIで設定可です<br>
※IPMessengerはデフォ設定でインストールしてください<br>

## ReplaceShortcut

ファイルサーバのお引越しだのでよくブッちぎれるショートカットを修復するためのブツです<br>
余裕こいて即興で書き始めると大体ミスって死にさらすんでテンプレでも用意しておこうかという<br>

## ScheduleShutdown

なんのこともないshutdownコマンド投げるだけのブツです<br>
予約キャンセルのコマンドってどーも忘れんのよねー<br>
はともかく画面ロック中とかはシャットダウンしませんので<BR>
PowerToysのAwakeで画面消灯とかロックとか抑止しておくと良いです<BR>

## MakePowerShellBatch

``MakePowerShellBatch.bat``はカレントの``MakePowerShellBatch.ps1``を起こすだけです<br>
``MakePowerShellBatch.ps1``はフォーム起動してD&Dで指定した``.ps1``をUTF-8(BOM付)に変換して起動用バッチを吐きます<BR>
GUI/CUI的なPowerShell用のバッチ生成とついでにPowerShell ISE起動するバッチ生成を選択できます<BR>

## MakePowerShellLink

``MakePowerShellLink.bat``はカレントの``MakePowerShellLink.ps1``を起こすだけです<br>
``MakePowerShellLink.ps1``はフォーム起動してD&Dで指定した``.ps1``の起動用ショートカットファイルを生成します<br>

空白を含むパスに使うと駄目な場合は大人しく``MakePowerShellBatch``のほーをどーぞ<br>

## AutoInstall

WingetのDSCとかいう機能で色々インストールします<BR>
なるほどコレが構成管理って奴なのかとゆー興味本位<BR>