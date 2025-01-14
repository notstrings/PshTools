Add-Type -AssemblyName "Microsoft.VisualBasic"

$ErrorActionPreference = "Stop"

$WSH = New-Object -ComObject WScript.Shell
$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

## 本体 #######################################################################

# 文字置換(大文字小文字を区別しない)
function local:ReplaceIgnoreCase([string] $text, [string] $from, [string] $to) {
    return [Microsoft.VisualBasic.Strings]::Replace($text, $from, $to, 1, -1, [Microsoft.VisualBasic.CompareMethod]::Text)
}

# 文字置換(大文字小文字を区別しない)※マップ
function local:ReplaceIgnoreCaseMap([string] $text, [PSCustomObject] $map) {
    # 字数の多い要素から置換
    $keys = $PathMap.Keys | Sort-Object @{Expression={$_.Length}; Ascending=$false}
    foreach ($key in $keys) {
        $text = ReplaceIgnoreCase $text -replace $key, $PathMap[$key]
    }
    return $text
}

# ファイル
function local:CleanupShortcutFile([System.IO.FileInfo] $Target) {
    CleanupShortcut $Target.FullName
}

# フォルダ
function local:CleanupShortcutDir([System.IO.DirectoryInfo] $Target) {
    $null = Write-Host "検査`t""$($Target.FullName)"""
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -Directory)) {
        CleanupShortcutDir $elm
    }
    foreach ($elm in @(Get-ChildItem -LiteralPath $Target.FullName -File -Filter "*.lnk")) {
        CleanupShortcutFile $elm
    }
}

# 本体
function local:CleanupShortcut([string] $TargetName) {
    # ショートカットファイルの内容を取得
    try {
        # 長大なパスの場合ショートカットファイルを開けないケースがある
        # ※環境依存のため手動修正が必要
        if ($TargetName.Length -lt 240) {
            $lnk = $WSH.createShortcut($TargetName)
        } else {
            $d = [System.IO.Path]::GetDirectoryName($TargetName)
            $f = [System.IO.Path]::GetFileName($TargetName)
            Set-Location -LiteralPath $d
            $lnk = $WSH.createShortcut($f)
        }
    } finally {
        Set-Location -LiteralPath $PSScriptRoot
    }
    $srcpath = $lnk.TargetPath
    $srcwork = $lnk.WorkingDirectory
    # 不要ショートカットファイル
    if ( [System.IO.Path]::GetDirectoryName($TargetName) -eq $srcpath ) {
        $null = Write-Host "削除`t代替元不要(自フォルダへのリダイレクト)`t""$TargetName""" -ForegroundColor Yellow
        Remove-Item $TargetName
        return
    }
    # 不正ショートカットファイル
    if ( ($srcpath -eq "") -and ($srcwork -ne "") ) {
        $null = Write-Host "修正`t代替元不正(ターゲット無/ワーキング有)`t""$TargetName""" -ForegroundColor Yellow
        # 修正後に処理を継続
        $lnk.TargetPath = $srcwork
        $lnk.WorkingDirectory = $srcpath
        $lnk.Save()
        $srcpath = $lnk.TargetPath
        $srcwork = $lnk.WorkingDirectory
    }
    # リンク切れ確認
    $srcexist = $true
    if ( $srcpath ) { $srcexist = $srcexist -and (Test-Path -LiteralPath $srcpath) }
    if ( $srcwork ) { $srcexist = $srcexist -and (Test-Path -LiteralPath $srcwork) }
    if ( $srcexist -eq $true ) {
        $null = Write-Host "無視`t処理不要`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    # スキップ対象チェック
    if ( (isSkipPath $srcpath) -or (isSkipPath $srcwork) ) {
        $null = Write-Host "無視`t無視リストに該当`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    # 代替候補パス生成
    $dstpath = ReplacePath $srcpath
    $dstwork = ReplacePath $srcwork
    # 代替候補パス確認
    if ( ($srcpath -eq $dstpath) -and ($srcwork -eq $dstwork) ) {
        $null = Write-Host "失敗`t代替先不変`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    if ( ($dstpath -eq "") -and ($dstwork -eq "") ) {
        $null = Write-Host "失敗`t代替先不正(ターゲット無/ワーキング無)`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    if ( ($dstpath -eq "") -and ($dstwork -ne "") ) {
        $null = Write-Host "失敗`t代替先不正(ターゲット無/ワーキング有)`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    # 代替候補パス不在
    $dstexist = $true
    if ($dstpath) { $dstexist = $dstexist -and (Test-Path -LiteralPath $dstpath) }
    if ($dstwork) { $dstexist = $dstexist -and (Test-Path -LiteralPath $dstwork) }
    if ( ($dstexist -eq $false)  ) {
        $null = Write-Host "失敗`t代替先不在`t""$TargetName""" -ForegroundColor Yellow
        return
    }
    # 入替
    $lnk.TargetPath = $dstpath
    $lnk.WorkingDirectory = $dstwork
    $lnk.Save()
    $null = Write-Host "成功`t""$TargetName""" -ForegroundColor Yellow
}

# 無視判定
function local:isSkipPath([string] $Path) {
    $skip = @(
        '\\192.168.1.10\*',     # 無視したいリンクを指定
        '\\ignore_serv000\*'    # 無視したいリンクを指定
    )
    foreach ($elm in $skip) {
        if ($path.ToUpper() -like $elm.ToUpper()) {
            return $true
        }
    }
    return $false
}

# 代替試行
function local:ReplacePath([string] $FilePath) {
    $fname = $FilePath
    # サーバやドライブ変更
    $fname = ReplaceIgnoreCase $fname '\\sample_serv0001' '\\sample_serv0002'
    $fname = ReplaceIgnoreCase $fname '\\sample_serv0002' '\\sample_serv0003'
    $fname = ReplaceIgnoreCase $fname '\\sample_serv0003' '\\sample_serv0004'
    $fname = $fname -replace '^([A-Z]):', '\\sample_serv0004\BACKUP\$1'
    # 単純なマッピング
    # ※ここでは文字列長は勝手に判断するので気にしなくていい
    $map = @{
        "\\serv0001\xxx\yyyy" = "\\serv0001\zzz\yyyy";
        "\\serv0002\xxx\yyyy" = "\\serv0001\zzz\yyyy";
    }
    $fname = ReplaceIgnoreCaseMap $fname $map
    # より複雑な制御
    $fname = $fname -replace '^([A-Z]):\\XXXXX_([^\\]+)', '$1\XXXXX\$2'
    return $fname
}

###############################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー")

try {
    $null = Write-Host "---$Title---"
	# 引数確認
    if ($args.Count -eq 0) {
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                CleanupShortcutDir (Get-Item $arg)
            } else {
                CleanupShortcutFile (Get-Item $arg)
            }
        }
    }
}
catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
