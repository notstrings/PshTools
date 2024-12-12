. "$($PSScriptRoot)/ModuleMisc.ps1"

$result = ShowDDDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" 
if ($result[0] -eq "OK") {
    Write-Host "選択されたファイル:"
    foreach ($file in $result[1]) {
        Write-Host "  - $file"
    }
} else {
    Write-Host "キャンセルされました。"
}
