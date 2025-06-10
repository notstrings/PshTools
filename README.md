# PshTools

日常的なファイル操作やドキュメント処理、画像変換、アーカイブ操作などを効率化するためのPowerShellスクリプト集です。
基本的に処理対象を同名WindowsバッチにD&Dして使用します。
内部的にはScoop等で導入したCUIツール（ImageMagick, Ghostscript, qpdf, 7zip, PDFsharp等）を活用しますので、使用には事前準備が必要です。

## 主な機能

- **ファイル・フォルダ操作**
    - 日付付きコピー（DateCopyF.ps1, DateCopyR.ps1）
    - 重複ファイルの検出・削除（RemoveDupFile.ps1）
    - ゴミ箱移動や進捗付きコピー（ModuleMisc.ps1）

- **PDF処理**
    - PDFへのテキスト追記（PDFText.ps1）
    - PDFへの日付印スタンプ（PDFDateStamp.ps1）
    - PDF画像変換・圧縮・平坦化など（ManipImage.ps1）

- **画像処理**
    - 画像のリサイズ・トリミング・傾き補正・注釈付記（ManipImage.ps1）

- **アーカイブ操作**
    - 7zipによる圧縮・展開・分割・暗号化（ManipArchive.ps1）

- **差分比較**
    - 画像・Wordファイルの差分比較（DiffImage.ps1, DiffWord.ps1）

- **フォルダ監視・スケジューラ**
    - フォルダ変更監視と通知（FolderMonitor.ps1）
    - 時刻指定でのスクリプト自動実行（ScheduleKicker.ps1）

- **バッチ・ショートカット作成**
    - PowerShellスクリプトからバッチファイルやショートカット自動生成（MakeBatch.ps1, MakeShortcut.ps1）

- **各種ユーティリティ**
    - メッセージボックスやファイル選択ダイアログ、トースト通知、IP Messenger連携、メール送信など（ModuleMisc.ps1）

## 使い方

1. **事前準備**  
   PowerShellの実行ポリシーを変更し、Scoopで必要なツールをインストールします。

   ```powershell
   Set-ExecutionPolicy RemoteSigned -scope CurrentUser
   iex (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
   scoop bucket add extras
   scoop install imagemagick
   scoop install ghostscript
   scoop install qpdf
   scoop install 7zip
   ```

2. **初回セットアップ**  
   PDF日付印など一部機能は追加セットアップが必要です。  
   それぞれ(例えばPDFDateStamp.ps1) の `Setup` 関数を実行してください。

3. **利用方法**  
   各`.bat` ファイルにファイルやフォルダをドラッグ＆ドロップすることで、対応する処理が実行されます。  
   設定ファイル（JSON）は初回実行時に自動生成され、GUIで編集可能です。

## 補足

- 詳細な各機能の使い方やパラメータは、各スクリプト先頭や関数コメントに記載されています。
- 拡張・カスタマイズも容易にできるよう、共通関数は ModuleMisc.ps1 に集約されています。

### 参考
- https://deepwiki.com/notstrings/PshTools
