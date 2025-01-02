$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$result = ShowFileListDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -FileFilter "\.txt$" 
if ($result[0] -eq "OK") {
    foreach ($file in $result[1]) {
        Write-Host $file
    }
}