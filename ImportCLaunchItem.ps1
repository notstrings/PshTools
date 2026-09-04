$ErrorActionPreference = "Stop"

## CLaunch用INI構造
## [xxx]
## [xxx]
## [Page000]
##   [Btn000]      // 重複する
##   [Btn001]      // 重複する
##   [ScrollStat]  // 重複する
## [Page001]
##   [Btn000]      // 重複する
##   [Btn001]      // 重複する
##   [ScrollStat]  // 重複する
## [SubMenus]
## [xxx]

# Iniファイルを読み込む
function ReadINI([string] $Content) {
    $Ret =  [ordered]@{}
    $section = "Undefined"
    foreach ($line in $Content -split "`r?`n") {
        if ($line -match '^\[(.*)\]$') {
            if(-not $Ret.Contains($matches[1].Trim())){
                $section = $matches[1].Trim()
                $Ret[$section] = [ordered]@{}
            }
        }
        if ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $Ret[$section][$key] = $val
        }
    }
    return $Ret
}

# Iniファイルを書き込む
function WriteINI([string] $Path, $Ini) {
    $lines = @()
    foreach ($section in $Ini.Keys) {
        $lines += "[$section]"
        foreach ($key in $Ini[$section].Keys) {
            $lines += "$key=$($Ini[$section][$key])"
        }
        $lines += ""
    }
    return $lines
}

# 環境変数に置換する
function ReplaceENV([string] $text){
    $text = $text -replace [regex]::Escape($env:SystemRoot),  "%SystemRoot%"
    $text = $text -replace [regex]::Escape($env:UserProfile), "%UserProfile%"
    $text = $text -replace [regex]::Escape($env:WinDir),      "%WinDir%"
    $text = $text -replace ".*\\Users\\[^\\]+\\scoop",        "%UserProfile%\scoop"
    return $text
}

# CLaunch用INIファイルの読み込み
function ReadCLaunchINI([string] $Path) {
    $Ini = [ordered]@{}
    # 読込
    $Content = Get-Content -Path $Path -Raw
    if($Content -eq ""){
        return $Ini
    }
    ## 先頭
    $Head = [regex]::Match($Content, "(.*?)\[Page\d+\]", 'Singleline').Groups[1].Value.Trim()
    $Ini["Head"] = ReadINI $Head
    ## 本体
    $Foot = [regex]::Match($Content, "(\[SubMenus\].*)", 'Singleline').Groups[1].Value.Trim()
    $Ini["Foot"] = ReadINI $Foot
    ## 末尾
    foreach ($pidx in (0..([int]$Ini["Head"]["Pages"]["Count"] - 1))) {
        $Body = [regex]::Match($Content, "(\[Page$(($pidx+0).ToString("000"))\].*?)(?:\[Page.*\]|(?=\[(?!Btn|Scr).*\])|$)", 'Singleline').Groups[1].Value.Trim()
        $Ini["Body"] += @(ReadINI $Body)
    }
    return $Ini
}

# CLaunch用INIファイルの書き込み
function WriteCLaunchINI([string] $Path, $Ini) {
    $lines = @()
    ## 先頭
    $lines += WriteINI -Path $Path -Ini $Ini["Head"]
    ## 本体
    foreach ($page in $Ini["Body"]) {
        $lines += WriteINI -Path $Path -Ini $page
    }
    ## 末尾
    $lines += WriteINI -Path $Path -Ini $Ini["Foot"]
    # 書込
    $lines -join "`r`n" | Out-File -FilePath $Path -Encoding Unicode
}

# CLaunch用INI置換
function CombineCLaunchINI($SrcINI, $DstINI) {
    # 固定設定
    $DstINI["Head"]["General"]["Status"]  = 00000000
    $DstINI["Head"]["General"]["Option1"] = 00205120
    $DstINI["Head"]["General"]["BackupFile"] = "C:\usr\srze\bin\cl64\Data\Backup.ini"

    # 置換辞書作成
    $List = @{}
    foreach ($page in $SrcINI["Body"]) {
        foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
            $key = $page[$btn]["Name"]
            $val = $page[$btn]
            $List[$key] = $val
        }
    }

    # 置換実行
    do{
        $cont = $false
        foreach ($page in $DstINI["Body"]) {
            foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
                $name = $page[$btn]["Name"]
                if( $List.ContainsKey($name) ) {
                    # 置換辞書でのリプレース
                    # ・アイテム位置は適用元のまま保持
                    if ($List[$name].Contains("Position")) {
                        $List[$name]["Position"] = $page[$btn]["Position"] # Ver.4.20以前
                    }
                    if ($List[$name].Contains("ViewMode1Pos") -and $List[$name].Contains("ViewMode2Pos")) {
                        $List[$name]["ViewMode1Pos"] = $page[$btn]["ViewMode1Pos"] # Ver.4.20以降
                        $List[$name]["ViewMode2Pos"] = $page[$btn]["ViewMode2Pos"] # Ver.4.20以降
                    }
                    if ($List[$name].Contains("Position") -and $List[$name].Contains("ViewMode1Pos") -and $List[$name].Contains("ViewMode2Pos")) {
                        $List[$name].Remove("Position")
                    }
                    $page[$btn] = $List[$name]
                    $List.Remove($name)
                    $cont = $true
                    # 再度
                    if ($cont) { break }
                }
            }
            # 再度
            if ($cont) { break }
        }
    }while($cont)

    # 余剰要素追加
    if ($List.Count -gt 0) {
        $pidx = ([int]$DstINI["Head"]["Pages"]["Count"])
        $DstINI["Body"] += [ordered]@{}
        $pname = "Page$($pidx.ToString("000"))"
        $DstINI["Body"][$pidx][$pname] = [ordered]@{
            Name        = "追加項目"
            ScrollMode1 = 0
            ScrollMode2 = 0
            Flag        = 00000000
            Count       = 0
        }
        $bidx = 0
        foreach ($name in $List.Keys) {
            $bkey = "Btn$($bidx.ToString("000"))"
            $DstINI["Body"][$pidx][$bkey] = $List[$name]
            $DstINI["Body"][$pidx][$bkey]["ViewMode1Pos"] = "$($bidx % 10),$([int]($bidx / 10))"
            $DstINI["Body"][$pidx][$bkey]["ViewMode2Pos"] = "$($bidx % 10),$([int]($bidx / 10))"
            $bidx = $bidx + 1
        }
        $DstINI["Body"][$pidx][$pname]["Count"] = $bidx
        $DstINI["Head"]["Pages"]["Count"] = $pidx + 1
    }
}

function RebuildCLaunchINI($Ini) {

    # ランチャアイテムのパス類に環境変数を使用させる
    foreach ($page in $Ini["Body"]) {
        foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
            $page[$btn]["File"]      = ReplaceENV $page[$btn]["File"]
            $page[$btn]["Directory"] = ReplaceENV $page[$btn]["Directory"]
            if ($page[$btn].Contains("IconFile")){
                $page[$btn]["IconFile"] = ReplaceENV $page[$btn]["IconFile"]
            }
        }
    }

    # ランチャアイテムの座標重複登録への対策
    foreach ($page in $Ini["Body"]) {
        # グリッド範囲取得
        $GridUsed = @{}
        $GridMaxX = 0
        $GridMaxY = 0
        foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
            $pos = $page[$btn]["ViewMode1Pos"]
            if ([string]::IsNullOrWhiteSpace($pos)) {
                $pos = "0,0"
            }
            $GridUsed[$pos] = $true
        }
        foreach ($elm in $GridUsed.Keys) {
            $x = [int]($elm.Split(",")[0])
            $y = [int]($elm.Split(",")[1])
            if ($x -gt $GridMaxX) { $GridMaxX = $x }
            if ($y -gt $GridMaxY) { $GridMaxY = $y } # 未使用/対称性のために保持
        }
        $GridMaxX = $GridMaxX + 1
        $GridMaxY = $GridMaxY + 1
        # 重複座標修正
        $UsedCrnt = @{}
        foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
            $pos = $page[$btn]["ViewMode1Pos"]
            if ([string]::IsNullOrWhiteSpace($pos)) {
                $pos = "0,0"
            }
            if (-not $UsedCrnt.ContainsKey($pos)) {
                $UsedCrnt[$pos] = $true
                continue
            }

            # 最初の空き座標を探す
            $idx = 0
            do {
                $x = $idx % $GridMaxX
                $y = [int]($idx / $GridMaxX)
                $newPos = "$x,$y"
                $idx++
            } while ($GridUsed.ContainsKey($newPos))

            # 新しい位置を設定
            $page[$btn]["ViewMode1Pos"] = $newPos

            # 使用済みに追加
            $UsedCrnt[$newPos] = $true
            $GridUsed[$newPos] = $true
        }
    }

    # モード2座標は行方不明しがちなのでモード1座標を統一
    foreach ($page in $Ini["Body"]) {
        foreach ($btn in $page.Keys | Where-Object { $_ -like "Btn*" }) {
            $page[$btn]["ViewMode2Pos"] = $page[$btn]["ViewMode1Pos"]
        }
    }

}

# CLaunch用デザイン設定の再構築※全ページをページデフォルトに統一
function RebuildCLDesignINI([string] $Path, [int]$PageCount) {
    $Ini = [ordered]@{}
    # 読込
    $Content = Get-Content -Path $Path -Raw
    if($Content -eq ""){
        return $Ini
    }
    ## 先頭読込
    $Head = [regex]::Match($Content, "(.*)\[PageDefault\]", 'Singleline').Groups[1].Value.Trim()
    ## 本体生成
    $Body = [regex]::Match($Content, "(\[PageDefault\].*?)(?:\[Page.*\]|$)", 'Singleline').Groups[1].Value.Trim()
    $Ini["PageDefault"] = ReadINI $Body
    $Ini["Page"] = @()
    foreach ($pidx in 0..($PageCount - 1)) {
        $pname = "Page$($pidx.ToString("000"))"
        $Ini["Page"] += [ordered]@{}
        $Ini["Page"][$pidx][$pname] = [ordered]@{}
        $Ini["Page"][$pidx]["TabN"] = $Ini["PageDefault"]["TabN"]
        $Ini["Page"][$pidx]["TabH"] = $Ini["PageDefault"]["TabH"]
        $Ini["Page"][$pidx]["TabS"] = $Ini["PageDefault"]["TabS"]
        $Ini["Page"][$pidx]["ButtonN"] = $Ini["PageDefault"]["ButtonN"]
        $Ini["Page"][$pidx]["ButtonH"] = $Ini["PageDefault"]["ButtonH"]
        $Ini["Page"][$pidx]["ButtonD"] = $Ini["PageDefault"]["ButtonD"]
    }
    ## 出力
    $lines = @()
    $lines += ""
    $lines += WriteINI -Path $Path -Ini $Ini["PageDefault"]
    foreach ($page in $Ini["Page"]) {
        $lines += WriteINI -Path $Path -Ini $page
    }
    $Head               | Set-Content -Path $Path -Encoding Unicode
    $lines -join "`r`n" | Add-Content -Path $Path -Encoding Unicode
}

# ランチャを停止
Stop-Process -Name "CLaunch" -Force -ErrorAction SilentlyContinue

# ランチャ設定合成
if (-not (Test-Path "C:\usr\srze\bin\cl64\Data\CLaunch.ini")) {
    Copy-Item "C:\usr\srze\bin\cl64\Data\CLaunch.org" "C:\usr\srze\bin\cl64\Data\CLaunch.ini"
} else {
    $SrcPath = "C:\usr\srze\bin\cl64\Data\CLaunch.org"
    $DstPath = "C:\usr\srze\bin\cl64\Data\CLaunch.ini"
    $SrcINI = ReadCLaunchINI $SrcPath
    $DstINI = ReadCLaunchINI $DstPath
    CombineCLaunchINI $SrcINI $DstINI
    RebuildCLaunchINI $DstINI
    WriteCLaunchINI $DstPath $DstINI
}

# デザイン設定再構築
if (Test-Path "C:\usr\srze\bin\cl64\Data\Design.ini") {
    RebuildCLDesignINI "C:\usr\srze\bin\cl64\Data\Design.ini" $DstINI["Head"]["Pages"]["Count"]
}

# アイコンキャッシュ削除
if (Test-Path "C:\usr\srze\bin\cl64\Data\ClIcons.bin") {
    Remove-Item "C:\usr\srze\bin\cl64\Data\ClIcons.bin"
}

# ランチャを再起動
Start-Process "C:\usr\srze\bin\cl64\CLaunch.exe"
