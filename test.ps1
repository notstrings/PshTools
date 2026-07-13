$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$result = ShowFileListDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -FileFilter ".*" -Files $args
if ($result.Result) {
    foreach ($file in $result.Files) {
        Write-Host $file
    }
}
