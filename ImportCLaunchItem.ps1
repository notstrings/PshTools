$ErrorActionPreference = "Stop"

# ちょっとした文字列置換処理かと思いきや派手にパースして大立ち回りする羽目にｗ

# CLaunch用INIファイルの読み込み
## ページ構成については特殊定義
function ReadCLaunchINI([string] $Path) {
    $Ini = @{}

    # コンテンツ取得(全体)
    $Content = Get-Content -Path $Path -Raw
    if($Content -eq ""){
        return $Ini
    }

    # 先頭部分
    $Head = [regex]::Match($Content, "(.*?)\[Page\d+\]", 'Singleline').Groups[1].Value.Trim()
    $section = "Undefined"
    foreach ($line in $Head -split "`r?`n") {
        if ($line -match '^\[(.*)\]$') {
            if(-not $Ini.ContainsKey($matches[1].Trim())){
                $section = $matches[1].Trim()
                $Ini[$section] = @{}
            }
        }
        if ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $Ini[$section][$key] = $val
        }
    }

    # 末尾部分
    $Foot = [regex]::Match($Content, "(\[SubMenus\].*)", 'Singleline').Groups[1].Value.Trim()
    $section = "Undefined"
    foreach ($line in $Foot -split "`r?`n") {
        if ($line -match '^\[(.*)\]$') {
            if(-not $Ini.ContainsKey($matches[1].Trim())){
                $section = $matches[1].Trim()
                $Ini[$section] = @{}
            }
        }
        if ($line -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $Ini[$section][$key] = $val
        }
    }

    # 特殊定義部分
    ## [Page000]
    ##   [Btn000]
    ##   [Btn001]
    ## [Page001]
    ##   [Btn000]
    ##   [Btn001]
    $Body = @{}
    foreach ($pidx in (0..([int]$Ini["Pages"]["Count"] - 1))) {
        # コンテンツ取得(ページ単位)
        $SttPageRegx = "(?:\[Page$(($pidx+0).ToString("000"))\])"
        $EndPageRegx = "(?:\[Page$(($pidx+1).ToString("000"))\]|(?=\[(?!Btn).*\])|$)"
        $PageContent = [regex]::Match($Content, "$SttPageRegx(.*?)$EndPageRegx", 'Singleline').Groups[1].Value.Trim()
        if($PageContent -eq ""){
            break
        }
        # ページ
        $Page = @{}
        foreach ($line in $PageContent -split "`r?`n") {
            if ($line -notmatch '^(.*?)=(.*)$') {
                break
            }
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $Page[$key] = $val
        }
        $Buttons = @{}
        foreach ($bidx in (0..([int]$Page["Count"] - 1))) {
            # コンテンツ取得(ボタン単位)
            $SttButtonRegx = "(?:\[Btn$(($bidx+0).ToString("000"))\])"
            $EndButtonRegx = "(?:\[Btn$(($bidx+1).ToString("000"))\]|$)"
            $ButtonContent = [regex]::Match($PageContent, "$SttButtonRegx(.*?)$EndButtonRegx", 'Singleline').Groups[1].Value.Trim()
            if($ButtonContent -eq ""){
                break
            }
            # ボタン
            $Items = @{}
            foreach ($line in $ButtonContent -split "`r?`n") {
                if ($line -notmatch '^(.*?)=(.*)$') {
                    break
                }
                $key = $matches[1].Trim()
                $val = $matches[2].Trim()
                $Items[$key] = $val
            }
            $Buttons["Btn$($bidx.ToString("000"))"] = $Items
        }
        $Page["Buttons"] = $Buttons

        $Body["Page$($pidx.ToString("000"))"] = $Page
    }
    $Ini["Body"] = $Body

    return $Ini
}

# CLaunch用INIファイルの書き込み
function WriteCLaunchINI([string] $Path, $Ini) {
    $lines = @()

    ## 先頭部分
    $Keys = @(
        "General",
        "Search",
        "Unhook",
        "Suppression",
        "SubMenu",
        "Keyboard",
        "HotKey",
        "HotKey000",
        "Mouse",
        "Edge",
        "Action",
        "Circle0",
        "Circle1",
        "AutoBoot000",
        "AutoBoot001",
        "AutoBoot002",
        "AutoBoot003",
        "Plugins",
        "Pages"
    )
    # foreach ($section in ($Ini.Keys | Sort-Object)) {
    foreach ($section in $Keys) {
        if ($section -ne "" -and $section -ne "Body") {
            $lines += "[$section]"
            foreach ($key in $Ini[$section].Keys) {
                $lines += "$key=$($Ini[$section][$key])"
            }
            $lines += ""
        }
    }

    # 特殊定義部分
    foreach ($page in ($Ini["Body"].Keys | Sort-Object)) {
        $lines += "[$page]"
        foreach ($key in ($Ini["Body"][$page].Keys | Sort-Object)) {
            $lines += "$key=$($Ini["Body"][$page][$key])"
        }
        foreach ($btn in ($Ini["Body"][$page]["Buttons"].Keys | Sort-Object)) {
            $lines += "[$btn]"
            foreach ($key in ($Ini["Body"][$page]["Buttons"][$btn].Keys | Sort-Object)) {
                $lines += "$key=$($Ini["Body"][$page]["Buttons"][$btn][$key])"
            }
        }
        $lines += ""
    }

    ## 末尾部分
    $Keys = @(
        "SubMenus"
    )
    # foreach ($section in ($Ini.Keys | Sort-Object)) {
    foreach ($section in $Keys) {
        if ($section -ne "" -and $section -ne "Body") {
            $lines += "[$section]"
            foreach ($key in $Ini[$section].Keys) {
                $lines += "$key=$($Ini[$section][$key])"
            }
            $lines += ""
        }
    }

    # 保存
    $lines -join "`r`n" | Out-File -FilePath $Path -Encoding Unicode
}

function ReplaceENV([string] $text){
    $text = $text -replace [regex]::Escape($env:SystemRoot),  "%SystemRoot%"
    $text = $text -replace [regex]::Escape($env:UserProfile), "%UserProfile%"
    $text = $text -replace [regex]::Escape($env:WinDir),      "%WinDir%"
    $text = $text -replace ".*\\Users\\[^\\]+\\scoop",        "%UserProfile%\scoop"
    return $text
}

# CLaunch用INI置換
function ReplaceCLaunchINI($SrcINI, $DstINI) {
    # 固定設定
    $DstINI["General"]["Status"]  = 00000000
    $DstINI["General"]["Option1"] = 00205120
    # 置換リスト作成
    $List = @{}
    foreach ($page in $SrcINI["Body"].Keys) {
        foreach ($btn in $SrcINI["Body"][$page]["Buttons"].Keys) {
            $key = $SrcINI["Body"][$page]["Buttons"][$btn]["Name"]
            $val = $SrcINI["Body"][$page]["Buttons"][$btn]
            $List[$key] = $val
        }
    }
    # 置換実行
    do{
        $cont = $false
        foreach ($page in $DstINI["Body"].Keys) {
            foreach ($btn in $DstINI["Body"][$page]["Buttons"].Keys) {
                $name = $DstINI["Body"][$page]["Buttons"][$btn]["Name"]
                if( $List.ContainsKey($name) ) {
                    # 置換リストでのリプレース
                    $DstINI["Body"][$page]["Buttons"][$btn]["Type"]       = $List[$name]["Type"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["File"]       = $List[$name]["File"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Parameter"]  = $List[$name]["Parameter"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Directory"]  = $List[$name]["Directory"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["WindowStat"] = $List[$name]["WindowStat"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Flag"]       = $List[$name]["Flag"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Tip"]        = $List[$name]["Tip"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Keyboard"]   = $List[$name]["Keyboard"]
                    if($List[$name].ContainsKey("IconFile")){
                        $DstINI["Body"][$page]["Buttons"][$btn]["IconFile"]  = $List[$name]["IconFile"]
                        $DstINI["Body"][$page]["Buttons"][$btn]["IconIndex"] = $List[$name]["IconIndex"]
                    }
                    $List.Remove($name)
                    $cont = $true
                    # 環境変数へのリプレース
                    $DstINI["Body"][$page]["Buttons"][$btn]["File"]      = ReplaceENV $DstINI["Body"][$page]["Buttons"][$btn]["File"]
                    $DstINI["Body"][$page]["Buttons"][$btn]["Directory"] = ReplaceENV $DstINI["Body"][$page]["Buttons"][$btn]["Directory"]
                    if ($DstINI["Body"][$page]["Buttons"][$btn].ContainsKey("IconFile")){
                        $DstINI["Body"][$page]["Buttons"][$btn]["IconFile"] = ReplaceENV $DstINI["Body"][$page]["Buttons"][$btn]["IconFile"]
                    }
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
        $pidx = ([int]$DstINI["Pages"]["Count"]) + 1
        $pkey = "Page$($pidx.ToString("000"))"
        $DstINI["Body"][$pkey] = @{
            Name        = "追加項目"
            ScrollMode1 = 0
            ScrollMode2 = 0
            Flag        = 00000000
            Count       = 0
            Buttons     = @{}
        }
        $bidx = 0
        $bkey = "Btn$($bidx.ToString("000"))"
        foreach ($name in $List.Keys) {
            $DstINI["Body"][$pkey]["Buttons"][$bkey] = @{
                Position   = $bidx
                Type       = $List[$name]["Type"]
                Name       = $List[$name]["Name"]
                File       = $List[$name]["File"]
                Parameter  = $List[$name]["Parameter"]
                Directory  = $List[$name]["Directory"]
                WindowStat = $List[$name]["WindowStat"]
                Flag       = $List[$name]["Flag"]
                Tip        = $List[$name]["Tip"]
                Keyboard   = $List[$name]["Keyboard"]
            }
            if($List[$name].ContainsKey("IconFile")){
                $DstINI["Body"][$pkey]["Buttons"][$bkey]["IconFile"]  = $List[$name]["IconFile"]
                $DstINI["Body"][$pkey]["Buttons"][$bkey]["IconIndex"] = $List[$name]["IconIndex"]
            }
            $bidx = $bidx + 1
            $bkey = "Btn$($bidx.ToString("000"))"
        }
        $DstINI["Body"][$pkey]["Count"] = $bidx
        $DstINI["Pages"]["Count"] = $pidx
    }
}

# ランチャを停止
Stop-Process -Name "CLaunch" -Force -ErrorAction SilentlyContinue

# 初期設定
if (-not (Test-Path "C:\usr\test\bin\cl64\Data\CLaunch.ini")) {
    Copy-Item "C:\usr\test\bin\cl64\Data\CLaunch.org" "C:\usr\test\bin\cl64\Data\CLaunch.ini"
} else {
    $SrcPath = "C:\usr\test\bin\cl64\Data\CLaunch.org"
    $DstPath = "C:\usr\test\bin\cl64\Data\CLaunch.ini"
    $SrcINI = ReadCLaunchINI $SrcPath
    $DstINI = ReadCLaunchINI $DstPath
    ReplaceCLaunchINI $SrcINI $DstINI
    WriteCLaunchINI $DstPath $DstINI
}

# ランチャを再起動
Start-Process "C:\usr\test\bin\cl64\CLaunch.exe"
