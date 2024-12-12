## ############################################################################
## とりあえずの関数置き場

Add-Type -AssemblyName "Microsoft.VisualBasic"
Add-type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

## ############################################################################
## 文字関連

<#
.SYNOPSIS
    文字列内の空白を正規化する
.DESCRIPTION
    文字列内の複数空白を1つして先頭と末尾の空白を削除します
.PARAMETER Text
    対象文字列
.EXAMPLE
    RestrictTextBlank "  あいう  えお  "
    結果:"あいう えお"
#>
function RestrictTextBlank {
    param (
        [Parameter(Mandatory = $false)] [string] $Text
    )
    begin {}
    process {
        $Text = [regex]::Replace($Text, "\s+", " ")   # 複数空白
        $Text = [regex]::Replace($Text, "^\s+", "")   # 先頭空白削除
        $Text = [regex]::Replace($Text, "\s+$", "")   # 末尾空白削除
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    全角・半角文字を変換する
.DESCRIPTION
    文字列中の全角英数/全角括弧"（）［］｛｝"を半角、半角カタカナを全角に変換します
.PARAMETER Text
    対象文字列
.EXAMPLE
    RestrictTextZenHan "Ｈｅｌｌｏ，ｗｏｒｌｄ！１２３（）［］｛｝　ｱｲｳｴｵ"
    結果:"Hello, world！123()[]{} アイウエオ"
#>
function RestrictTextZenHan() {
    param (
        [Parameter(Mandatory = $false)] [string] $Text
    )
    begin {}
    process {
        $Text = [regex]::Replace($Text, "[Ａ-Ｚａ-ｚ０-９　（）［］｛｝]+",{ 
            param($match)
            return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Narrow)
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        $Text = [regex]::Replace($Text, "[ｦ-ﾟ]+",{ 
            param($match)
            return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Wide)
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    ファイル・フォルダ名の日付部分を正規化します
.DESCRIPTION
    入力された文字列中の日付部分を指定されたフォーマットに正規化します。西暦/和暦/2桁年数に対応しています。
.PARAMETER Text
    対象文字列
.PARAMETER Format
    出力する日付のフォーマット (例: yyyyMMdd, dd-MM-yyyy)
.PARAMETER RefDate
    参考日時 (省略可能)
.EXAMPLE
    RestrictDate "2023年12月25日_報告書.docx" "yyyyMMdd"
    結果:"20231225_報告書.docx"
.NOTES
    * 変換可能な日付形式はYYYY-MM-DD/YYYY.MM.DD/YYYY年MM月DD日/和暦YY-MM-DD/和暦YY.MM.DD/和暦YY年MM月DD日です
    * 年号省略の場合は表記年度と参照年度が一致する場合だけ処理します
#>
function RestrictTextDate {
    param (
        [parameter(Mandatory=$false)] [string]$Text,
        [parameter(Mandatory=$true)]  [string]$Format,
        [parameter(Mandatory=$false)] [datetime]$RefDate
    )
    begin {}
    process {
        # 日本のカレンダー情報を取得
        $info = New-Object CultureInfo("ja-jp", $true)
        $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
        ## YYYY-MM-DD or YYYY.MM.DD
        $Text = [regex]::Replace($Text, "(?<![0-9]+)(19|20)(\d\d)([.-])([1-9]|0[1-9]|1[0-2])(\3)([1-9]|0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
            param($match)
            $name = $match.Value.ToUpper()
            $name = $name.Replace(".","-")
            $date = [DateTime]::ParseExact($name, "yyyy-M-d", $null) 
            if($date){ return $date.ToString($Format) }else{ return $match.Value }
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        ## YYYY年MM月DD日
        $Text = [regex]::Replace($Text, "(?<![0-9]+)(19|20)(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
            param($match)
            $name = $match.Value.ToUpper()
            $date = [DateTime]::ParseExact($name, "yyyy年M月d日", $null) 
            if($date){ return $date.ToString($Format) }else{ return $match.Value }
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        ## 和暦YY-MM-DD or 和暦YY.MM.DD 
        $Text = [regex]::Replace($Text, "(令和|\bR|平成|\bH|昭和|\bS|明治|\bM|大正|\bT)(\d{1,2})([.-])([1-9]|0[1-9]|1[0-2])(\3)([1-9]|0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
            param($match)
            $name = $match.Value.ToUpper()
            $name = $name.Replace(".","-")
            $name = $name.Replace("R","令和")
            $name = $name.Replace("H","平成")
            $name = $name.Replace("S","昭和")
            $name = $name.Replace("M","明治")
            $name = $name.Replace("T","大正")
            $date = [DateTime]::ParseExact($name, "gy-M-d", $info) 
            if($date){ return $date.ToString($Format) }else{ return $match.Value }
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        ## 和暦YY年MM月DD日
        $Text = [regex]::Replace($Text, "(令和|\bR|平成|\bH|昭和|\bS|明治|\bM|大正|\bT)(\d{1,2}|元)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
            param($match)
            $name = $match.Value.ToUpper()
            $name = $name.Replace("R","令和")
            $name = $name.Replace("H","平成")
            $name = $name.Replace("S","昭和")
            $name = $name.Replace("M","明治")
            $name = $name.Replace("T","大正")
            $date = [DateTime]::ParseExact($name, "gy年M月d日", $info) 
            if($date){ return $date.ToString($Format) }else{ return $match.Value }
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        # 年号省略の場合は表記年度と参照年度が一致する場合だけ処理
        if ($RefDate -ne $null) {
            ## YY-MM-DD or YY.MM.DD
            $Text = [regex]::Replace($Text, "\b(\d\d)([.-])(0[1-9]|1[0-2])(\2)(0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
                param($match)
                $name = $match.Value.ToUpper()
                $name = $name.Replace(".","-")
                $nameyy = ($RefDate.Year).ToString().Substring(0,2) + $name
                $dateyy = [DateTime]::ParseExact($nameyy, "yyyy-M-d", $null) 
                $namegg = $RefDate.ToString("ggg", $info) + $name
                $dategg = [DateTime]::ParseExact($namegg, "gggy-M-d", $info) 
                if( ($dateyy) -and ($RefDate.Year -eq $dateyy.Year) ){
                    return $dateyy.ToString($Format)
                }elseif( ($dategg) -and ($RefDate.Year -eq $dategg.Year) ){
                    return $dategg.ToString($Format)
                }else{
                    return $match.Value
                }
            }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
            ## YY年MM月DD日
            $Text = [regex]::Replace($Text, "\b(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
                param($match)
                $name = $match.Value.ToUpper()
                $name = $name.Replace(".","-")
                $nameyy = ($RefDate.Year).ToString().Substring(0,2) + $name
                $dateyy = [DateTime]::ParseExact($nameyy, "yyyy年M月d日", $null) 
                $namegg = $RefDate.ToString("ggg", $info) + $name
                $dategg = [DateTime]::ParseExact($namegg, "gy年M月d日", $info) 
                if( ($dateyy) -and ($RefDate.Year -eq $dateyy.Year) ){
                    return $dateyy.ToString($Format)
                }elseif( ($dategg) -and ($RefDate.Year -eq $dategg.Year) ){
                    return $dategg.ToString($Format)
                }else{
                    return $match.Value
                }
            }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        }
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    ローマ字をひらがなにします
.DESCRIPTION
    ローマ字をひらがなにしようとしますが
    ローマ字表記方式は死ぬほど色々あるので完璧ではありません
.PARAMETER Text
    対象文字列
.EXAMPLE
    TryConvertRoma2Kana "aiueo"
    結果:"あいうえお"
#>
function TryConvertRoma2Kana {
    param (
        [Parameter(Mandatory = $true)] [string] $Text
    )
    begin {}
    process {
        $RomajiMapA = @{
            "Â"= "Aー"; "Î"= "Iー"; "Ûー"= "U"; "Ê"= "Eー"; "Ô"= "Oー"; 
            "Ā"= "Aー"; "Ī"= "Iー"; "Ūー"= "U"; "Ē"= "Eー"; "Ō"= "Oー"; 
            "nn"= "ん"; 
            "qa"= "っq"; "qi"= "っq"; "qu"= "っq"; "qe"= "っq"; "qo"= "っq";
            "kk"= "っk"; "ss"= "っs"; "tt"= "っt"; "qn"= "っn"; "hh"= "っh"; 
            "mm"= "っm"; "yy"= "っy"; "rr"= "っr"; "ww"= "っw"; "gg"= "っg"; 
            "zz"= "っz"; "dd"= "っd"; "bb"= "っb"; "pp"= "っp"; "tc"= "っc"; 
            "ff"= "っf"; "jj"= "っj";
        }
        $RomajiMapB = @{
            "kya"= "きゃ"; "kyi"= "きぃ"; "kyu"= "きゅ"; "kye"= "きぇ"; "kyo"= "きょ";
            "sha"= "しゃ"; "shi"= "し";   "shu"= "しゅ"; "she"= "しぇ"; "sho"= "しょ";
            "sya"= "しゃ"; "syi"= "しぃ"; "syu"= "しゅ"; "sye"= "しぇ"; "syo"= "しょ";
            "cha"= "ちゃ"; "chi"= "ち";   "chu"= "ちゅ"; "che"= "ちぇ"; "cho"= "ちょ";
            "tya"= "ちゃ"; "tyi"= "ちぃ"; "tyu"= "ちゅ"; "tye"= "ちぇ"; "tyo"= "ちょ";
            "nya"= "にゃ"; "nyi"= "にぃ"; "nyu"= "にゅ"; "nye"= "にぇ"; "nyo"= "にょ";
            "hya"= "ひゃ"; "hyi"= "ひぃ"; "hyu"= "ひゅ"; "hye"= "ひぇ"; "hyo"= "ひょ";
            "mya"= "みゃ"; "myi"= "みぃ"; "myu"= "みゅ"; "mye"= "みぇ"; "myo"= "みょ";
            "rya"= "りゃ"; "ryi"= "りぃ"; "ryu"= "りゅ"; "rye"= "りぇ"; "ryo"= "りょ";
            "gya"= "ぎゃ"; "gyi"= "ぎぃ"; "gyu"= "ぎゅ"; "gye"= "ぎぇ"; "gyo"= "ぎょ";
            "zya"= "じゃ"; "zyi"= "じぃ"; "zyu"= "じゅ"; "zye"= "じぇ"; "zyo"= "じょ";
            "dya"= "ぢゃ"; "dyi"= "ぢぃ"; "dyu"= "ぢゅ"; "dye"= "ぢぇ"; "dyo"= "ぢょ";
            "bya"= "びゃ"; "byi"= "びぃ"; "byu"= "びゅ"; "bye"= "びぇ"; "byo"= "びょ";
            "pya"= "ぴゃ"; "pyi"= "ぴぃ"; "pyu"= "ぴゅ"; "pye"= "ぴぇ"; "pyo"= "ぴょ";
            "kwa"= "くぁ"; "kwi"= "くぃ"; "kwu"= "くゅ"; "kwe"= "くぇ"; "kwo"= "くぉ";
            "swa"= "すぁ"; "swi"= "すぃ"; "swu"= "すゅ"; "swe"= "すぇ"; "swo"= "すぉ";
            "twa"= "つぁ"; "twi"= "つぃ"; "twu"= "つゅ"; "twe"= "つぇ"; "two"= "つぉ";
            "nwa"= "ぬぁ"; "nwi"= "ぬぃ"; "nwu"= "ぬゅ"; "nwe"= "ぬぇ"; "nwo"= "ぬぉ";
            "hwa"= "ふぁ"; "hwi"= "ふぃ"; "hwu"= "ふゅ"; "hwe"= "ふぇ"; "hwo"= "ふぉ";
            "mwa"= "むぁ"; "mwi"= "むぃ"; "mwu"= "むゅ"; "mwe"= "むぇ"; "mwo"= "むぉ";
            "rwa"= "るぁ"; "rwi"= "るぃ"; "rwu"= "るゅ"; "rwe"= "るぇ"; "rwo"= "るぉ";
            "gwa"= "ぐぁ"; "gwi"= "ぐぃ"; "gwu"= "ぐゅ"; "gwe"= "ぐぇ"; "gwo"= "ぐぉ";
            "zwa"= "ずぁ"; "zwi"= "ずぃ"; "zwu"= "ずゅ"; "zwe"= "ずぇ"; "zwo"= "ずぉ";
            "bwa"= "ぶぁ"; "bwi"= "ぶぃ"; "bwu"= "ぶゅ"; "bwe"= "ぶぇ"; "bwo"= "ぶぉ";
            "pwa"= "ぷぁ"; "pwi"= "ぷぃ"; "pwu"= "ぷゅ"; "pwe"= "ぷぇ"; "pwo"= "ぷぉ";
            "tja"= "てゃ"; "tji"= "てぃ"; "tju"= "てゅ"; "tje"= "てぇ"; "tjo"= "てぉ";
            "dja"= "でゃ"; "dji"= "でぃ"; "dju"= "でゅ"; "dje"= "でぇ"; "djo"= "でぉ";
            "tva"= "とぁ"; "tvi"= "とぃ"; "tvu"= "とゅ"; "tve"= "とぇ"; "tvo"= "とぉ";
            "dva"= "どぁ"; "dvi"= "どぃ"; "dvu"= "どゅ"; "dve"= "どぇ"; "dvo"= "どぉ";
            "ka"= "か"; "ki"= "き"; "ku"= "く"; "ke"= "け"; "ko"= "こ";
            "sa"= "さ"; "si"= "し"; "su"= "す"; "se"= "せ"; "so"= "そ";
            "ta"= "た"; "ti"= "ち"; "tu"= "つ"; "te"= "て"; "to"= "と";
            "na"= "な"; "ni"= "に"; "nu"= "ぬ"; "ne"= "ね"; "no"= "の";
            "ha"= "は"; "hi"= "ひ"; "hu"= "ふ"; "he"= "へ"; "ho"= "ほ";
            "ma"= "ま"; "mi"= "み"; "mu"= "む"; "me"= "め"; "mo"= "も";
            "ya"= "や"; "yi"= "ゐ"; "yu"= "ゆ"; "ye"= "ゑ"; "yo"= "よ";
            "ra"= "ら"; "ri"= "り"; "ru"= "る"; "re"= "れ"; "ro"= "ろ";
            "wa"= "わ"; "wi"= "ゐ"; "wu"= "ぅ"; "we"= "ゑ"; "wo"= "を";
            "ga"= "が"; "gi"= "ぎ"; "gu"= "ぐ"; "ge"= "げ"; "go"= "ご";
            "za"= "ざ"; "zi"= "じ"; "zu"= "ず"; "ze"= "ぜ"; "zo"= "ぞ";
            "da"= "だ"; "di"= "ぢ"; "du"= "づ"; "de"= "で"; "do"= "ど";
            "ba"= "ば"; "bi"= "び"; "bu"= "ぶ"; "be"= "べ"; "bo"= "ぼ";
            "pa"= "ぱ"; "pi"= "ぴ"; "pu"= "ぷ"; "pe"= "ぺ"; "po"= "ぽ";
            "va"= "ゔぁ"; "vi"= "ゔぃ"; "vu"= "ゔ";   "ve"= "ゔぇ"; "vo"= "ゔぉ";
            "ja"= "じゃ"; "ji"= "じ";   "ju"= "じゅ"; "je"= "じぇ"; "jo"= "じょ";
            "xa"= "ぁ"; "xi"= "ぅ"; "xu"= "ぃ"; "xe"= "ぇ"; "xo"= "ぉ";
            "a"= "あ"; "i"= "い"; "u"= "う"; "e"= "え"; "o"= "お"; 
            "n"= "ん" 
        }
        $keys = $RomajiMapA.Keys | Sort-Object @{Expression={$_.Length}; Ascending=$false}
        foreach ($key in $keys) {
            $value = $RomajiMapA[$key]
            $Text = $Text -replace $key, $value
        }
        $keys = $RomajiMapB.Keys | Sort-Object @{Expression={$_.Length}; Ascending=$false}
        foreach ($key in $keys) {
            $value = $RomajiMapB[$key]
            $Text = $Text -replace $key, $value
        }
        return $ret
    }
    end {}        
}

<#
.SYNOPSIS
    文字列から括弧とその中身を削除します
.DESCRIPTION
    指定された文字列から括弧 `(...)` とその中身をすべて削除します
    入れ子になった括弧にも対応しています
    括弧の片割れが残った場合先頭or末尾に相方がある前提で削除します
.PARAMETER Text
    括弧を削除する文字列
.EXAMPLE
    RemoveAllBrackets -Text "(example) text (with) brackets"
    "example"及び"with" を返します
#>
Function RemoveAllBrackets {
    param (
        [Parameter(Mandatory = $true)] [string] $Text
    )
    begin {}
    process {
        $buff = $Text
        do {
            $buff = [regex]::Replace($buff, "\([^\(]*?\)","")
        } until (
            $buff -eq [regex]::Replace($buff, "\([^\(]*?\)","")
        )
        $buff = [regex]::Replace($buff, ".*\)", "")
        $buff = [regex]::Replace($buff, "\(.*", "")
        return $buff
    }
    end {}
}

<#
.SYNOPSIS
    テキストファイルのエンコーディングを判定します
.DESCRIPTION
    指定されたテキストファイルのエンコーディングをBOMの有無や文字コードの出現頻度に基づいて判定します
    判定対象は日本語のテキストファイル(Shift-JIS/EUC-JP/UTF-8)だけですので過信しないように
.PARAMETER Path
    エンコーディングを判定するテキストファイルのパスを指定します
.EXAMPLE
    $encoding = GetEncodingSimple "C:\temp\sample.txt"
    Write-Host "エンコーディング: $($encoding.EncodingName)"
.NOTES
   * BOMが存在する場合はBOMに基づいてエンコーディングを判定します
   * BOMが存在しない場合は日本語文字コードの出現頻度を基にShift-JIS、EUC-JP、UTF-8 のいずれかを判定します
   * BOMが存在しない場合は上記以外の可能性を考慮しません(まぁ普通にテキスト弄ってる分には十分でしょ...?)
#>
Function GetEncodingSimple {
    param (
        [Parameter(Mandatory = $true)] [string] $Path
    )
    begin {}
    process {
        $TxtData = [System.IO.File]::ReadAllBytes($Path)
        # BOMによるエンコード判定
        foreach ($elm in [System.Text.Encoding]::GetEncodings()) {
            $oAmb = $elm.GetEncoding().GetPreamble()
            if ($oAmb.Length -gt 0 ) {
                if (Compare-Object $TxtData[0..($oAmb.Length - 1)] $oAmb -IncludeEqual -ExcludeDifferent) {
                    return $elm.GetEncoding()
                }
            }
        }
        # BOM無しなので日本語前提にSJIS/EUC/UTF8の何れかをソレっぽさで判定する
        # 元ネタは特徴的な文字コードを数え上げてスコアリングしてたが...
        # ・スコアリングせずソレっぽい位置を数えるように変更
        # ・末尾のややこしい処理は移植面倒なのでバッサリ破棄
        # 元々の処理が推定だし単発スクリプトならコレで良いでしょ多分
        # なお元ネタは以下URLのC#コード
        # ・http://dobon.net/vb/dotnet/string/detectcode.html
        $SJISCount = 0..($TxtData.Length - 2) | Where-Object {
            $b1 = $TxtData[$_ + 0]
            $b2 = $TxtData[$_ + 1]
            ((0x81 -le $b1 -and $b1 -le 0x9F) -or (0xE0 -le $b1 -and $b1 -le 0xFC)) -and
            ((0x40 -le $b2 -and $b2 -le 0x7E) -or (0x80 -le $b2 -and $b2 -le 0xFC))
        }
        $EUCJCount = 0..($TxtData.Length - 2) | Where-Object {
            $b1 = $TxtData[$_ + 0]
            $b2 = $TxtData[$_ + 1]
            (((0xA1 -le $b1 -and $b1 -le 0xFE) -and (0xA1 -le $b2 -and $b2 -le 0xFE)) -or
            ((0x8E -eq $b1)                   -and (0xA1 -le $b2 -and $b2 -le 0xDF)))
        }
        $UTF8Count = 0..($TxtData.Length - 2) | Where-Object {
            $b1 = $TxtData[$_ + 0]
            $b2 = $TxtData[$_ + 1]
            (0xC0 -le $b1 -and $b1 -le 0xDF) -and (0x80 -le $b2 -and $b2 -le 0xBF)
        }
        if ($SJISCount.Count -gt $EUCJCount.Count -and $SJISCount.Count -gt $UTF8Count.Count) {
            return [System.Text.Encoding]::GetEncoding("Shift-JIS")
        } elseif ($EUCJCount.Count -gt $SJISCount.Count -and $EUCJCount.Count -gt $UTF8Count.Count) {
            return [System.Text.Encoding]::GetEncoding("EUC-JP")
        } else {
            return [System.Text.Encoding]::GetEncoding("UTF-8")
        }
    }
    end {}
}

## ############################################################################
## UI関連

<#
.SYNOPSIS
    ファイル選択ダイアログを表示し選択されたファイルのパスを返します
.DESCRIPTION
    Windows標準のファイル選択ダイアログを表示しユーザーが選択したファイルのパスを配列形式で返します
    タイトル/フィルター/初期ディレクトリ/複数選択の可否を選択できます
.PARAMETER Title
   ダイアログのタイトルを指定します
.PARAMETER Filter
   表示するファイルの種類のフィルターを指定します
.PARAMETER InitialDirectory
   最初に表示するディレクトリを指定します
.PARAMETER Multiselect
   複数のファイルを選択できるようにするかどうかを指定します。$trueで複数選択可能。
.EXAMPLE
#>
function ShowFileDialog {
    param (
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $false)] [string] $Filter  = "テキストファイル(*.txt)|*.txt",
        [Parameter(Mandatory = $false)] [string] $InitialDirectory = $PSScriptRoot,
        [Parameter(Mandatory = $false)] [string] $Multiselect = $false
    )
    begin {}
    process {
        $FDlg = New-Object System.Windows.Forms.OpenFileDialog
        $FDlg.Title            = $Title
        $FDlg.InitialDirectory = $InitialDirectory
        $FDlg.Filter           = $Filter
        $FDlg.Multiselect      = $Multiselect
        $null = $FDlg.ShowDialog()
        return $FDlg.FileNames
    }
    end {}
}

<#
.SYNOPSIS
    ドラッグ＆ドロップで受け取ったファイルを選択するためのダイアログを表示します
.DESCRIPTION
    FileDDBox 関数はタイトル/メッセージ/ファイルフィルター/ボタンラベルを指定して
    ドラッグ＆ドロップで受け取ったファイルを表示するダイアログボックスを作成します
    ユーザーがボタンを押すと選択したボタンのラベルと選択されたファイルのパスが返されます
.PARAMETER Title
    ダイアログボックスのタイトルに設定する文字列です
.PARAMETER Message
    ダイアログボックスに表示するメッセージの文字列です
.PARAMETER Filter
    ドラッグ＆ドロップで受け付けるファイル名のフィルター(正規表現)です
.PARAMETER ButtonA
    ボタンAのラベルに設定する文字列です
.PARAMETER ButtonB
    ボタンBのラベルに設定する文字列です
.EXAMPLE
    # ファイルを選択し選択したボタンのラベルとファイルパスを表示します
    $result = ShowDDDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -Filter "\.txt$" -ButtonA "OK" -ButtonB "キャンセル"
    if ($result[0] -eq "OK") {
        Write-Host "選択されたファイル:"
        foreach ($file in $result[1]) {
            Write-Host "  - $file"
        }
    } else {
        Write-Host "キャンセルされました。"
    }
#>
function ShowDDDialog {
    param (
        [Parameter(Mandatory = $true)]  [string] $Title,
        [Parameter(Mandatory = $true)]  [string] $Message,
        [Parameter(Mandatory = $false)] [string] $Filter  = ".*",
        [Parameter(Mandatory = $false)] [string] $ButtonA = "OK",
        [Parameter(Mandatory = $false)] [string] $ButtonB = "Cancel"
    )
    begin {}
    process {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title                                     # タイトル
        $form.Size = New-Object System.Drawing.Size(300,200)    # ウィンドウサイズ
        $form.StartPosition = 'CenterScreen'                    # 表示位置
        $form.Topmost = $true                                   # TopMost
        $tableLayoutPanel1 = New-Object System.Windows.Forms.TableLayoutPanel
        $panel1  = New-Object System.Windows.Forms.TableLayoutPanel
        $panel2  = New-Object System.Windows.Forms.Panel
        $label   = New-Object System.Windows.Forms.Label
        $listbox = New-Object System.Windows.Forms.ListBox
        $button1 = New-Object System.Windows.Forms.Button
        $button2 = New-Object System.Windows.Forms.Button
        $tableLayoutPanel1.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tableLayoutPanel1.RowCount = 2
        $null = $tableLayoutPanel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tableLayoutPanel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tableLayoutPanel1.Controls.Add($panel1, 0, 0)
        $null = $tableLayoutPanel1.Controls.Add($panel2, 0, 1)
        $null = $form.Controls.Add($tableLayoutPanel1)
        $panel1.Dock = [System.Windows.Forms.DockStyle]::Fill
        $panel1.RowCount = 2
        $null = $panel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle))
        $null = $panel1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)))
        $null = $panel1.Controls.Add($label, 0, 0)
        $null = $panel1.Controls.Add($listbox, 0, 1)
        $null = $panel2.Controls.Add($button2)
        $null = $panel2.Controls.Add($button1)
        $panel2.Dock = [System.Windows.Forms.DockStyle]::Fill
        $label.Dock = [System.Windows.Forms.DockStyle]::Fill
        $label.Text = $Message
        $label.AutoSize = $true
        $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $listbox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $listbox.AllowDrop = $True
        $null = $listbox.Add_DragEnter({
            $_.Effect = "All"
        })
        $null = $listbox.Add_DragDrop({
            foreach($elm in @($_.Data.GetData("FileDrop"))) {
                if( [System.IO.Path]::GetFileName($elm) -match $Filter ){
                    [void]$Listbox.Items.Add($elm)
                }
            }
        })
        $button1.Dock = [System.Windows.Forms.DockStyle]::Right
        $button1.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
        $button1.Text = $ButtonA
        $button1.UseVisualStyleBackColor = $true
        $null = $button1.Add_Click({
            $form.Text = $ButtonA
            $form.Close()
        })
        $button2.Dock = [System.Windows.Forms.DockStyle]::Right
        $button2.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
        $button2.Text = $ButtonB
        $button2.UseVisualStyleBackColor = $true
        $null = $button2.Add_Click({
            $form.Text = $ButtonB
            $form.Close()
        })
        $null = $form.ShowDialog()
        return @($form.Text, $listbox.Items)
    }
    end {}
}

## ############################################################################
## WebAPI関連

<#
.SYNOPSIS
    ファイルを重複しないファイル名にして移動します
.DESCRIPTION
    ファイルを重複しないファイル名にして移動します
.PARAMETER SrcPath
    移動元のファイルまたはフォルダのパス
.PARAMETER DstPath
    移動先のファイルまたはフォルダのパス
.EXAMPLE
    MoveItemWithUniqName -SrcPath "C:\Temp\test.txt" -DstPath "C:\Temp\test.txt" -isDir $false
    "C:\Temp\test.txt"を"C:\Temp\test.txt"に移動します
    "C:\Temp\test.txt"が既に存在する場合は"C:\Temp\test (1).txt"に移動します
#>
function MoveItemWithUniqName {
    param (
        [Parameter(Mandatory = $true)] [string] $SrcPath,
        [Parameter(Mandatory = $true)] [string] $DstPath
    )
    begin {}
    process {
        if ($SrcPath -ne $DstPath) {
            $sUniq = $DstPath
            $lUniq = 1
            while( (Test-Path -Path $sUniq) ) {
                if ((Get-Item $SrcPath).PSIsContainer) {
                    $dname = [System.IO.Path]::GetDirectoryName($DstPath)
                    $fname = [System.IO.Path]::GetFileName($DstPath)
                    $ename = ""
                }else{
                    $dname = [System.IO.Path]::GetDirectoryName($DstPath)
                    $fname = [System.IO.Path]::GetFileNameWithoutExtension($DstPath)
                    $ename = [System.IO.Path]::GetExtension($DstPath)
                }
                $sUniq = [System.IO.Path]::Combine($dname, "$fname ($lUniq)" + $ename)
                $lUniq++
            }
            $null = Move-Item -Path $SrcPath -Destination $sUniq -Force
        }
    }
    end {}
}

<#
.SYNOPSIS
    ファイルをゴミ箱に移動します
.DESCRIPTION
    指定されたファイルをゴミ箱に移動します
.PARAMETER Path
    ゴミ箱に移動するファイルのパス
.EXAMPLE
    MoveTrush -Path "C:\Temp\test.txt"
    "C:\Temp\test.txt"をゴミ箱に移動します
#>
function MoveTrush {
    param (
        [Parameter(Mandatory = $true)] [string] $Path
    )
    begin {}
    process {
        $dpath = [System.IO.Path]::GetDirectoryName($Path)
        $fpath = [System.IO.Path]::GetFileName($Path)
        $shell = New-Object -comobject Shell.Application
        $shell.Namespace($dpath).ParseName($fpath).InvokeVerb("delete")
    }
    end {}
}

## ############################################################################

<#
.SYNOPSIS
    テキストを翻訳します
.DESCRIPTION
    Google Translate APIを使用して指定されたテキストを翻訳します
.PARAMETER Text
    翻訳するテキスト
.PARAMETER SrcLang
    翻訳元の言語コード
.PARAMETER DstLang
    翻訳先の言語コード
.EXAMPLE
    GoogleTranslate -Text "Hello, world!" -SrcLang "en" -DstLang "ja"
    "Hello, world!" を英語から日本語に翻訳します
#>
function GoogleTranslate {
    param (
        [Parameter(Mandatory = $true)] [string] $Text,
        [Parameter(Mandatory = $true)] [string] $SrcLang,
        [Parameter(Mandatory = $true)] [string] $DstLang
    )
    begin {}
    process {
        $Uri = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=$($SrcLang)&tl=$($DstLang)&dt=t&q=$Text"
        $Res = Invoke-RestMethod -Uri $Uri -Method Get
        return $Res[0].SyncRoot | ForEach-Object { $_[0] }
    }
    end {}
}

<#
.SYNOPSIS
    ひらがなを漢字にします
.DESCRIPTION
    ひらがなをGoogleTranslate(※IMEの方)で漢字にします
    変換候補は全部最初のものになるため精度はあまり良くありません
.PARAMETER Text
    対象文字列
.EXAMPLE
    TryConvertKana2Kanji "きょうは"
    結果:"今日は"
#>
function GoogleIME {
    param (
        [Parameter(Mandatory = $true)] [string] $Text
    )
    begin {}
    process {
        $ret = ""
        $url = "http://www.google.com/transliterate?langpair=ja-Hira|ja&text={0}" -f ([System.Web.HttpUtility]::UrlEncode($Text))
        $res = [System.Net.HttpWebRequest]::Create($url).GetResponse()
        $rdr = New-Object System.IO.StreamReader($res.GetResponseStream())
        $content = $rdr.ReadToEnd()
        if ($content -ne "") {
            $jsonResponse = $content | ConvertFrom-Json
            foreach ($elm in $jsonResponse) {
                $ret += $elm[1][0]
            }
        }
        return $ret
    }
    end {}
}

# Export-ModuleMember -Function *
