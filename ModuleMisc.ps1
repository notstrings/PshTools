## ############################################################################
## とりあえずの関数置き場

Add-Type -AssemblyName "Microsoft.VisualBasic"
Add-type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

## ############################################################################
## オブジェクト操作

<#
.SYNOPSIS
    オブジェクトをディープコピー
.DESCRIPTION
    オブジェクトをディープコピー
.PARAMETER Text
    対象オブジェクト
.NOTES
    あまりちゃんと考えてないんで複雑な物はダメかも
#>
function DeepCopyObj {
    param (
        [Parameter(Mandatory = $true)] [object] $obj
    )
    begin {}
    process {
        $typ = $obj.GetType()
        $ret = New-Object -TypeName $typ.FullName
        foreach ($prop in $typ.GetProperties()){
            if ($prop.CanRead -and $prop.CanWrite) {
                $prop.SetValue($ret, $prop.GetValue($obj))
            }
        }
        return $ret
    }
    end {}
}

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
    全角文字を半角に変換します
.DESCRIPTION
    指定した全角文字を半角に変換します
.PARAMETER Text
    対象文字列
.EXAMPLE
    RestrictTextZen "Ｈｅｌｌｏ，ｗｏｒｌｄ！１２３　ｱｲｳｴｵ"
    結果:"Hello, world！123　ｱｲｳｴｵ"
#>
function RestrictTextZen() {
    param (
        [Parameter(Mandatory = $false)] [string] $Text,
        [Parameter(Mandatory = $false)] [string] $Chars = "Ａ-Ｚａ-ｚ０-９　（）［］｛｝"
    )
    begin {}
    process {
        $Text = [regex]::Replace($Text, "[$(Chars)]+",{ 
            param($match)
            return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Narrow)
        }, [system.text.regularexpressions.regexoptions]::IgnoreCase)
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    半角カタカナを全角に変換します
.DESCRIPTION
    半角カタカナを全角に変換します
.PARAMETER Text
    対象文字列
.EXAMPLE
    RestrictTextHan "Ｈｅｌｌｏ，ｗｏｒｌｄ！１２３　ｱｲｳｴｵ"
    結果:"Hello, world！123　アイウエオ"
#>
function RestrictTextHan() {
    param (
        [Parameter(Mandatory = $false)] [string] $Text
    )
    begin {}
    process {
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
            $Text = [regex]::Replace($Text, "(?<![0-9]+)(\d\d)([.-])(0[1-9]|1[0-2])(\2)(0[1-9]|[12][0-9]|3[01])(?![0-9]+)",{
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
            $Text = [regex]::Replace($Text, "(?<![0-9]+)(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
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
    つーか多分日本人でも完璧に書けるヤツ居ねぇんじゃね
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
            "ja" = "じゃ"; "ji" = "じ";   "ju" = "じゅ"; "je" = "じぇ"; "jo" = "じょ";
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
            "va" = "ゔぁ"; "vi" = "ゔぃ"; "vu" = "ゔ";   "ve" = "ゔぇ"; "vo" = "ゔぉ";
            "a" = "あ"; "i" = "い"; "u" = "う"; "e" = "え"; "o" = "お"; 
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
            "xa"= "ぁ"; "xi"= "ぃ"; "xu"= "ぅ"; "xe"= "ぇ"; "xo"= "ぉ";
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
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    ひらがなをローマ字にします
.DESCRIPTION
    ひらがなをローマ字にしようとしますが
    強引に逆変換しただけなので正式なものではありません
    ※ただし「すうぃーつ」とかいう元々日本語には無い綴りも許容します
.PARAMETER Text
    対象文字列
.EXAMPLE
    TryConvertKana2Roma "あいうえお"
    結果:"aiueo"
#>
function TryConvertKana2Roma {
    param (
        [Parameter(Mandatory = $true)] [string] $Text
    )
    begin {}
    process {
        $KanaMapA = @{
            "ん"= "nn"; 
            "っ"= "tt";
        }
        $KanaMapB = @{
            "きゃ"= "kya"; "きぃ"= "kyi"; "きゅ"= "kyu"; "きぇ"= "kye"; "きょ"= "kyo";
            "しゃ"= "sya"; "しぃ"= "syi"; "しゅ"= "syu"; "しぇ"= "sye"; "しょ"= "syo";
            "ちゃ"= "cha"; "ちぃ"= "tyi"; "ちゅ"= "chu"; "ちぇ"= "che"; "ちょ"= "cho";
            "にゃ"= "nya"; "にぃ"= "nyi"; "にゅ"= "nyu"; "にぇ"= "nye"; "にょ"= "nyo";
            "ひゃ"= "hya"; "ひぃ"= "hyi"; "ひゅ"= "hyu"; "ひぇ"= "hye"; "ひょ"= "hyo";
            "みゃ"= "mya"; "みぃ"= "myi"; "みゅ"= "myu"; "みぇ"= "mye"; "みょ"= "myo";
            "りゃ"= "rya"; "りぃ"= "ryi"; "りゅ"= "ryu"; "りぇ"= "rye"; "りょ"= "ryo";
            "ぎゃ"= "gya"; "ぎぃ"= "gyi"; "ぎゅ"= "gyu"; "ぎぇ"= "gye"; "ぎょ"= "gyo";
            "じゃ"= "ja";  "じぃ"= "zyi"; "じゅ"= "ju";  "じぇ"= "je";  "じょ"= "jo";
            "ぢゃ"= "dya"; "ぢぃ"= "dyi"; "ぢゅ"= "dyu"; "ぢぇ"= "dye"; "ぢょ"= "dyo";
            "びゃ"= "bya"; "びぃ"= "byi"; "びゅ"= "byu"; "びぇ"= "bye"; "びょ"= "byo";
            "ぴゃ"= "pya"; "ぴぃ"= "pyi"; "ぴゅ"= "pyu"; "ぴぇ"= "pye"; "ぴょ"= "pyo";
            "くぁ"= "kwa"; "くぃ"= "kwi"; "くゅ"= "kwu"; "くぇ"= "kwe"; "くぉ"= "kwo";
            "すぁ"= "swa"; "すぃ"= "swi"; "すゅ"= "swu"; "すぇ"= "swe"; "すぉ"= "swo";
            "つぁ"= "twa"; "つぃ"= "twi"; "つゅ"= "twu"; "つぇ"= "twe"; "つぉ"= "two";
            "ぬぁ"= "nwa"; "ぬぃ"= "nwi"; "ぬゅ"= "nwu"; "ぬぇ"= "nwe"; "ぬぉ"= "nwo";
            "ふぁ"= "hwa"; "ふぃ"= "hwi"; "ふゅ"= "hwu"; "ふぇ"= "hwe"; "ふぉ"= "hwo";
            "むぁ"= "mwa"; "むぃ"= "mwi"; "むゅ"= "mwu"; "むぇ"= "mwe"; "むぉ"= "mwo";
            "るぁ"= "rwa"; "るぃ"= "rwi"; "るゅ"= "rwu"; "るぇ"= "rwe"; "るぉ"= "rwo";
            "ぐぁ"= "gwa"; "ぐぃ"= "gwi"; "ぐゅ"= "gwu"; "ぐぇ"= "gwe"; "ぐぉ"= "gwo";
            "ずぁ"= "zwa"; "ずぃ"= "zwi"; "ずゅ"= "zwu"; "ずぇ"= "zwe"; "ずぉ"= "zwo";
            "ぶぁ"= "bwa"; "ぶぃ"= "bwi"; "ぶゅ"= "bwu"; "ぶぇ"= "bwe"; "ぶぉ"= "bwo";
            "ぷぁ"= "pwa"; "ぷぃ"= "pwi"; "ぷゅ"= "pwu"; "ぷぇ"= "pwe"; "ぷぉ"= "pwo";
            "てゃ"= "tja"; "てぃ"= "tji"; "てゅ"= "tju"; "てぇ"= "tje"; "てぉ"= "tjo";
            "でゃ"= "dja"; "でぃ"= "dji"; "でゅ"= "dju"; "でぇ"= "dje"; "でぉ"= "djo";
            "とぁ"= "tva"; "とぃ"= "tvi"; "とゅ"= "tvu"; "とぇ"= "tve"; "とぉ"= "tvo";
            "どぁ"= "dva"; "どぃ"= "dvi"; "どゅ"= "dvu"; "どぇ"= "dve"; "どぉ"= "dvo";
            "ゔぁ"= "va";  "ゔぃ"= "vi";  "ゔ"= "vu";    "ゔぇ"= "ve";  "ゔぉ"= "vo";
            "あ"= "a";  "い"= "i";  "う"= "u";  "え"= "e";  "お"= "o"; 
            "か"= "ka"; "き"= "ki"; "く"= "ku"; "け"= "ke"; "こ"= "ko";
            "さ"= "sa"; "し"= "si"; "す"= "su"; "せ"= "se"; "そ"= "so";
            "た"= "ta"; "ち"= "ti"; "つ"= "tu"; "て"= "te"; "と"= "to";
            "な"= "na"; "に"= "ni"; "ぬ"= "nu"; "ね"= "ne"; "の"= "no";
            "は"= "ha"; "ひ"= "hi"; "ふ"= "hu"; "へ"= "he"; "ほ"= "ho";
            "ま"= "ma"; "み"= "mi"; "む"= "mu"; "め"= "me"; "も"= "mo";
            "や"= "ya";             "ゆ"= "yu";             "よ"= "yo";
            "ら"= "ra"; "り"= "ri"; "る"= "ru"; "れ"= "re"; "ろ"= "ro";
            "わ"= "wa"; "ゐ"= "wi";             "ゑ"= "we"; "を"= "wo";
            "が"= "ga"; "ぎ"= "gi"; "ぐ"= "gu"; "げ"= "ge"; "ご"= "go";
            "ざ"= "za"; "じ"= "zi"; "ず"= "zu"; "ぜ"= "ze"; "ぞ"= "zo";
            "だ"= "da"; "ぢ"= "di"; "づ"= "du"; "で"= "de"; "ど"= "do";
            "ば"= "ba"; "び"= "bi"; "ぶ"= "bu"; "べ"= "be"; "ぼ"= "bo";
            "ぱ"= "pa"; "ぴ"= "pi"; "ぷ"= "pu"; "ぺ"= "pe"; "ぽ"= "po";
            "ぁ"= "xa"; "ぃ"= "xi"; "ぅ"= "xu"; "ぇ"= "xe"; "ぉ"= "xo";
        }
        $keys = $KanaMapA.Keys | Sort-Object @{Expression={$_.Length}; Ascending=$false}
        foreach ($key in $keys) {
            $value = $KanaMapA[$key]
            $Text = $Text -replace $key, $value
        }
        $keys = $KanaMapB.Keys | Sort-Object @{Expression={$_.Length}; Ascending=$false}
        foreach ($key in $keys) {
            $value = $KanaMapB[$key]
            $Text = $Text -replace $key, $value
        }
        $Text = $Text -replace "Aー", "Â" 
        $Text = $Text -replace "Iー", "Î" 
        $Text = $Text -replace "Uー", "U" 
        $Text = $Text -replace "Eー", "Ê" 
        $Text = $Text -replace "Oー", "Ô" 
        return $Text
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
   * 元ネタ:http://dobon.net/vb/dotnet/string/detectcode.html
   * BOMが存在する場合はBOMに基づいてエンコーディングを判定します
   * BOMが存在しない場合は日本語文字コードの出現頻度を基にShift-JIS、EUC-JP、UTF-8 のいずれかを判定します
   * BOMが存在しない場合は上記以外の可能性を考慮しません
   * コレでダメなら``https://github.com/hnx8/ReadJEnc``辺りをどーぞ
#>
Function AutoGuessEncodingSimple {
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

<#
.SYNOPSIS
    2つのテキストファイルの内容を比較し結果を取得します
.DESCRIPTION
    指定された2つのテキストファイルの内容を比較して
    両方のファイルに存在する行と左右片側にのみ存在する行を配列として返します
.PARAMETER LHSPath
    比較対象の左側のファイルのパスを指定します
.PARAMETER RHSPath
    比較対象の右側のファイルのパスを指定します
.PARAMETER Encoding
    ファイルのエンコーディングを指定します
.EXAMPLE
    $ret = DiffContent -LHSPath "file1.txt" -RHSPath "file2.txt" -Encoding ([System.Text.Encoding]::GetEncoding("Shift_JIS"))
    $ret[0] | ForEach-Object { Write-Host "Line" $_.ReadCount $_ } # 両方のファイルに存在する行
    $ret[1] | ForEach-Object { Write-Host "Line" $_.ReadCount $_ } # 左側のファイルにのみ存在する行の配列
    $ret[2] | ForEach-Object { Write-Host "Line" $_.ReadCount $_ } # 右側のファイルにのみ存在する行の配列
#>
function DiffContent {
    param (
        [Parameter(Mandatory=$true)]  [string]$LHSPath,
        [Parameter(Mandatory=$true)]  [string]$RHSPath,
        [Parameter(Mandatory=$false)] [string]$Encoding = "UTF8"
    )
    begin {}
    process {
        $Both = @()
        $LHSOnly = @()
        $RHSOnly = @()
        $LHS = @(Get-Content $LHSPath -Encoding $Encoding)
        $RHS = @(Get-Content $RHSPath -Encoding $Encoding)
        Compare-Object -ReferenceObject $LHS -DifferenceObject $RHS -IncludeEqual |
        ForEach-Object {
            if($_.SideIndicator -eq "<=") {
                $LHSOnly += $_.InputObject
            } elseif ($_.SideIndicator -eq "=>") {
                $RHSOnly += $_.InputObject
            } elseif ($_.SideIndicator -eq "==") {
                $Both += $_.InputObject
            }
        } | Out-Null
        return $Both, $LHSOnly, $RHSOnly
    }
    end {}
}

## ############################################################################
## UI関連

<#
.SYNOPSIS
    コンソールの親ウィンドウを取得
.DESCRIPTION
    コンソールの親ウィンドウを取得
#>
function GetConsoleWindow {
    param ()
    begin {}
    process {
        Add-Type -Name ConsoleAPI -Namespace Win32Util -MemberDefinition '
            [DllImport("Kernel32.dll")]
            public static extern IntPtr GetConsoleWindow();
        '
        return [Win32Util.ConsoleAPI]::GetConsoleWindow()
    }
    end {}
}

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
#>
function ShowFileDialog {
    param (
        [Parameter(Mandatory = $true)]  [string] $Title,
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
        $frmMain.Owner = [System.Windows.Forms.Form]::FromHandle((GetConsoleWindow)) # 無駄な足掻きをしておく
        $null = $frmMain.ShowDialog()
        return $FDlg.FileNames
    }
    end {}
}

<#
.SYNOPSIS
    フォルダ選択ダイアログを表示し選択されたフォルダのパスを返します
.DESCRIPTION
    Windows標準のフォルダ選択ダイアログを表示しユーザーが選択したフォルダのパスを配列形式で返します
    タイトル/フィルター/初期ディレクトリ/複数選択の可否を選択できます
.PARAMETER Description
    ダイアログの説明
.PARAMETER InitialDirectory
    最初に表示するディレクトリを指定します
#>
function ShowFolderDialog {
    param (
        [Parameter(Mandatory = $true)]  [string] $Description,
        [Parameter(Mandatory = $false)] [string] $InitialDirectory = $PSScriptRoot
    )
    begin {}
    process {
        $FDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $FDlg.Description      = $Description
        $FDlg.InitialDirectory = $InitialDirectory
        $frmMain.Owner = [System.Windows.Forms.Form]::FromHandle((GetConsoleWindow)) # 無駄な足掻きをしておく
        $null = $frmMain.ShowDialog()
        return $FDlg.SelectedPath
    }
    end {}
}

<#
.SYNOPSIS
    ドラッグ＆ドロップで受け取ったファイルを選択するためのダイアログを表示します
.DESCRIPTION
    タイトル/メッセージ/ファイルフィルター/初期リストを指定して
    ドラッグ＆ドロップで受け取ったファイルを表示するダイアログボックスを作成します
    ユーザーがOKボタンを押すと選択した結果と選択されたファイルのパスが返されます
.PARAMETER Title
    ダイアログボックスのタイトルに設定する文字列です
.PARAMETER Message
    ダイアログボックスに表示するメッセージの文字列です
.PARAMETER FileFilter
    ドラッグ＆ドロップで受け付けるファイル名のフィルター(正規表現)です
.PARAMETER FileList
    初期リスト
.EXAMPLE
    # ファイルを選択し選択したボタンのラベルとファイルパスを表示します
    $result = ShowFileListDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -FileFilter "\.txt$" -FileList @("aaa.txt","bbb.txt") 
    if ($result[0] -eq "OK") {
        foreach ($file in $result[1]) {
            Write-Host $file
        }
    }
#>
function ShowFileListDialog {
    param (
        [Parameter(Mandatory = $true)]  [string]   $Title,
        [Parameter(Mandatory = $true)]  [string]   $Message,
        [Parameter(Mandatory = $false)] [string]   $FileFilter  = ".*",
        [Parameter(Mandatory = $false)] [string[]] $FileList
    )
    begin {}
    process {
        # フォーム生成
        $frmMain = New-Object System.Windows.Forms.Form
        $frmMain.Text = $Title                                     # タイトル
        $frmMain.Size = New-Object System.Drawing.Size(480,320)    # ウィンドウサイズ
        $frmMain.StartPosition = 'CenterScreen'                    # 表示位置
        $frmMain.Topmost = $true                                   # TopMost
        $frmMain.Padding = New-Object System.Windows.Forms.Padding(5)

        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel
            $pnlBody = New-Object System.Windows.Forms.Panel
                $lblDD   = New-Object System.Windows.Forms.Label
                $lbxDD = New-Object System.Windows.Forms.ListBox
            $pnlTail = New-Object System.Windows.Forms.Panel
                $btnOK     = New-Object System.Windows.Forms.Button
                $btnCancel = New-Object System.Windows.Forms.Button
        
        $tlpMain.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tlpMain.RowCount = 2
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)
        
            $pnlBody.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlBody.Controls.Add($lbxDD)
            $null = $pnlBody.Controls.Add($lblDD)
        
                $lblDD.Dock = [System.Windows.Forms.DockStyle]::Top
                $lblDD.Text = $Message
                $lblDD.AutoSize = $true
                $lblDD.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

                $lbxDD.Dock = [System.Windows.Forms.DockStyle]::Fill
                $lbxDD.AllowDrop = $true
                $null = $lbxDD.Add_DragEnter({
                    $_.Effect = "All"
                })
                $null = $lbxDD.Add_DragDrop({
                    @($_.Data.GetData("FileDrop")) | ForEach-Object {
                        if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                            [void]$lbxDD.Items.Add($_)
                        }
                    }
                })

            $pnlTail.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlTail.Controls.Add($btnCancel)
            $null = $pnlTail.Controls.Add($btnOK)
        
                $btnOK.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnOK.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnOK.Text = "OK"
                $btnOK.UseVisualStyleBackColor = $true
                $null = $btnOK.Add_Click({
                    $frmMain.DialogResult = "OK"
                    $frmMain.Close()
                })
        
                $btnCancel.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnCancel.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnCancel.Text = "Cancel"
                $btnCancel.UseVisualStyleBackColor = $true
                $null = $btnCancel.Add_Click({
                    $frmMain.DialogResult = "Cancel"
                    $frmMain.Close()
                })

        # フォーム表示
        if ($null -ne $FileList) {
            $FileList | ForEach-Object {
                if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                    [void]$lbxDD.Items.Add($_)
                }
            }
        }
        $frmMain.DialogResult = "Cancel"
        $frmMain.AcceptButton = $btnOK
        $frmMain.CancelButton = $btnCancel
        $frmMain.Owner = [System.Windows.Forms.Form]::FromHandle((GetConsoleWindow)) # 無駄な足掻きをしておく
        $null = $frmMain.ShowDialog()
        return @($frmMain.DialogResult, $lbxDD.Items)
    }
    end {}
}

<#
.SYNOPSIS
    ドラッグ＆ドロップで受け取ったファイルを選択するためのダイアログを表示します
    ※オプションの指定が追加されています
.DESCRIPTION
    タイトル/メッセージ/ファイルフィルター/初期リストを指定して
    ドラッグ＆ドロップで受け取ったファイルを表示するダイアログボックスを作成します
    ユーザーがOKボタンを押すと選択した結果と選択されたファイルのパスが返されます
    ※オプションの指定が追加されています
.PARAMETER Title
    ダイアログボックスのタイトルに設定する文字列です
.PARAMETER Message
    ダイアログボックスに表示するメッセージの文字列です
.PARAMETER FileFilter
    ドラッグ＆ドロップで受け付けるファイル名のフィルター(正規表現)です
.PARAMETER FileList
    初期リスト
.PARAMETER Options
    オプションリスト
.EXAMPLE
    # ファイルを選択し選択したボタンのラベルとファイルパスを表示します
    $result = ShowFileListDialogWithOption -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -FileFilter "\.txt$" -FileList @("aaa.txt","bbb.txt") -Options @("aaa","bbb")
    if ($result[0] -eq "OK") {
        foreach ($file in $result[1]) {
            Write-Host $file $result[2]
        }
    }
#>
function ShowFileListDialogWithOption {
    param (
        [Parameter(Mandatory = $true)]  [string]   $Title,
        [Parameter(Mandatory = $true)]  [string]   $Message,
        [Parameter(Mandatory = $false)] [string]   $FileFilter  = ".*",
        [Parameter(Mandatory = $false)] [string[]] $FileList,
        [Parameter(Mandatory = $false)] [string[]] $Options
    )
    begin {}
    process {
        # フォーム生成
        $frmMain = New-Object System.Windows.Forms.Form
        $frmMain.Text = $Title                                     # タイトル
        $frmMain.Size = New-Object System.Drawing.Size(480,320)    # ウィンドウサイズ
        $frmMain.StartPosition = 'CenterScreen'                    # 表示位置
        $frmMain.Topmost = $true                                   # TopMost
        $frmMain.Padding = New-Object System.Windows.Forms.Padding(5)

        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel
            $pnlBody = New-Object System.Windows.Forms.Panel
                $lblDD   = New-Object System.Windows.Forms.Label
                $lbxDD = New-Object System.Windows.Forms.ListBox
                $grpOpt = New-Object System.Windows.Forms.GroupBox
                    $flpOpt = New-Object System.Windows.Forms.FlowLayoutPanel
            $pnlTail = New-Object System.Windows.Forms.Panel
                $btnOK     = New-Object System.Windows.Forms.Button
                $btnCancel = New-Object System.Windows.Forms.Button
        
        $tlpMain.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tlpMain.RowCount = 2
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)
        
            $pnlBody.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlBody.Controls.Add($grpOpt)
            $null = $pnlBody.Controls.Add($lbxDD)
            $null = $pnlBody.Controls.Add($lblDD)
        
                $lblDD.Dock = [System.Windows.Forms.DockStyle]::Top
                $lblDD.Text = $Message
                $lblDD.AutoSize = $true
                $lblDD.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

                $lbxDD.Dock = [System.Windows.Forms.DockStyle]::Fill
                $lbxDD.AllowDrop = $true
                $null = $lbxDD.Add_DragEnter({
                    $_.Effect = "All"
                })
                $null = $lbxDD.Add_DragDrop({
                    @($_.Data.GetData("FileDrop")) | ForEach-Object {
                        if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                            [void]$lbxDD.Items.Add($_)
                        }
                    }
                })

                $grpOpt.Dock = [System.Windows.Forms.DockStyle]::Bottom
                $grpOpt.Text = "Options"
                $grpOpt.Height = 50
                $null = $grpOpt.Controls.Add($flpOpt)
                    $flpOpt.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $Checked = $true
                    $Options | ForEach-Object {
                        $rdoOpt = New-Object System.Windows.Forms.RadioButton
                        $rdoOpt.Text = $_
                        $rdoOpt.Checked = $Checked
                        $rdoOpt.AutoSize = $true
                        $flpOpt.Controls.Add($rdoOpt)
                        $Checked = $false
                    }
        
            $pnlTail.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlTail.Controls.Add($btnCancel)
            $null = $pnlTail.Controls.Add($btnOK)
        
                $btnOK.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnOK.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnOK.Text = "OK"
                $btnOK.UseVisualStyleBackColor = $true
                $null = $btnOK.Add_Click({
                    $frmMain.DialogResult = "OK"
                    $frmMain.Close()
                })
        
                $btnCancel.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnCancel.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnCancel.Text = "Cancel"
                $btnCancel.UseVisualStyleBackColor = $true
                $null = $btnCancel.Add_Click({
                    $frmMain.DialogResult = "Cancel"
                    $frmMain.Close()
                })

        # フォーム表示
        if ($null -ne $FileList) {
            $FileList | ForEach-Object {
                if( [System.IO.Path]::GetFileName($_) -match $FileFilter ){
                    [void]$lbxDD.Items.Add($_)
                }
            }
        }
        $frmMain.DialogResult = "Cancel"
        $frmMain.AcceptButton = $btnOK
        $frmMain.CancelButton = $btnCancel
        $frmMain.Owner = [System.Windows.Forms.Form]::FromHandle((GetConsoleWindow)) # 無駄な足掻きをしておく
        $null = $frmMain.ShowDialog()
        return @($frmMain.DialogResult, $lbxDD.Items, ($flpOpt.Controls | Where-Object {$_.Checked -eq $true} | Select-Object -ExpandProperty Text))
    }
    end {}
}

<#
.SYNOPSIS
    プロパディグリッドを使った汎用設定ダイアログを表示します
.DESCRIPTION
    プロパディグリッドを使った汎用設定ダイアログを表示します
    引数値はディープコピーして使うので副作用は外部に伝搬しません
.PARAMETER Title
    ダイアログボックスのタイトルに設定する文字列です
.PARAMETER Setting
    クラスインスタンスを設定してください
    ※多分連想配列でも大丈夫です
.EXAMPLE
    class AppSettings {
        [System.ComponentModel.Description("名前")]
        [string]$AppName
        [int]$Version
        [bool]$AutoUpdate
        [string]$LogFilePath
        [System.Diagnostics.SourceLevels]$LogLevel
    }
    $settings = [AppSettings]@{
        AppName = "My Application"
        Version = 1
        AutoUpdate = $true
        LogFilePath = "C:\app.log"
        LogLevel = [System.Diagnostics.SourceLevels]::Information
    }
    $ret = ShowSettingDialog "Title" $settings
    if ($ret[0] -eq "OK") {
        $ret[1]
    }
#>
function ShowSettingDialog {
    param (
        [Parameter(Mandatory = $true)]  [string]          $Title,
        [Parameter(Mandatory = $true)]  [System.Object]   $Setting
    )
    begin {}
    process {
        # ディープコピー
        $ret = DeepCopyObj $Setting
  
        # フォーム生成
        $frmMain = New-Object System.Windows.Forms.Form
        $frmMain.Text = $Title                                     # タイトル
        $frmMain.Size = New-Object System.Drawing.Size(480,320)    # ウィンドウサイズ
        $frmMain.StartPosition = 'CenterParent'                    # 表示位置
        $frmMain.Padding = New-Object System.Windows.Forms.Padding(5)
  
        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel
            $pnlBody = New-Object System.Windows.Forms.Panel
                $grdProp = New-Object System.Windows.Forms.PropertyGrid
            $pnlTail = New-Object System.Windows.Forms.Panel
                $btnOK     = New-Object System.Windows.Forms.Button
                $btnCancel = New-Object System.Windows.Forms.Button
        
        $tlpMain.Dock = [System.Windows.Forms.DockStyle]::Fill
        $tlpMain.RowCount = 2
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)
        
            $pnlBody.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlBody.Controls.Add($grdProp)
        
                $grdProp.Dock = [System.Windows.Forms.DockStyle]::Fill
                $grdProp.SelectedObject = $ret
        
            $pnlTail.Dock = [System.Windows.Forms.DockStyle]::Fill
            $null = $pnlTail.Controls.Add($btnCancel)
            $null = $pnlTail.Controls.Add($btnOK)
        
                $btnOK.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnOK.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnOK.Text = "OK"
                $btnOK.UseVisualStyleBackColor = $true
                $null = $btnOK.Add_Click({
                    $frmMain.DialogResult = "OK"
                    $frmMain.Close()
                })
        
                $btnCancel.Dock = [System.Windows.Forms.DockStyle]::Right
                $btnCancel.Size = New-Object System.Drawing.Size(128, 36) # ボタン巾のみ指定可能
                $btnCancel.Text = "Cancel"
                $btnCancel.UseVisualStyleBackColor = $true
                $null = $btnCancel.Add_Click({
                    $frmMain.DialogResult = "Cancel"
                    $frmMain.Close()
                })
  
        # フォーム表示
        $frmMain.DialogResult = "Cancel"
        # $frmMain.AcceptButton = $btnOK
        # $frmMain.CancelButton = $btnCancel
        $frmMain.Owner = [System.Windows.Forms.Form]::FromHandle((GetConsoleWindow)) # 無駄な足掻きをしておく
        $null = $frmMain.ShowDialog()
  
        return @($frmMain.DialogResult, $ret)
    }
    end {}
  }

# タスクトレイ常駐用アイコン生成
# 元ネタ:https://aquasoftware.net/blog/?p=1244
function local:GenTaskTrayIcon([uint32] $ARGB) {
    # PowerShell(ぽい)アイコン画像バイナリ
    # ・16x16 1bit/pixelインデックスカラー画像
    # ・パレット色を書き換えてアイコン背景色を一括変更する
    $icon = 'AAABAAEAEBAQAAEABAB4AAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAEDAAAAJT1tIgAAAAZQTFRFAAB/////8DxOgwAAAC1JREFUCNdjYGBgYGFg4GNgYGdgYG5gYGxgYHgARcwHQIJyDAwWNQwGNQxgAACDjAYG7YuK+QAAAABJRU5ErkJggg=='
    $strm = New-Object System.IO.MemoryStream(,[System.Convert]::FromBase64String($icon))
    $strm.Seek(0x3f, [System.IO.SeekOrigin]::Begin) > $null
    $ARGB = $ARGB -band 0x00ffffff
    $strm.WriteByte($ARGB -shr 16 -band 0xff)
    $strm.WriteByte($ARGB -shr  8 -band 0xff)
    $strm.WriteByte($ARGB -shr  0 -band 0xff)
    $strm.Seek(0x0, [System.IO.SeekOrigin]::Begin) > $null
    return New-Object System.Drawing.Icon($strm)
}

<#
.SYNOPSIS
    タスクトレイに常駐して指定したスクリプトを定期的に実行します
.DESCRIPTION
    タスクトレイに常駐して指定したスクリプトを定期的に実行します
    指定スクリプトはタスクトレイアイコンを左クリックすることで任意タイミングで実行可能です
.PARAMETER Name
    タスクの名前
.PARAMETER Color
    アイコンの色
.PARAMETER Conf
    設定コードブロック(タスクトレイアイコン右クリックで起動)
.PARAMETER Exec
    実行コードブロック(タスクトレイアイコン右クリックorインターバルで起動)
.PARAMETER Interval
    実行インターバル
.NOTES
    元ネタ:https://aquasoftware.net/blog/?p=1244
#>
function RunInTray {
    param (
        [Parameter(Mandatory = $true)] [string]      $Name,
        [Parameter(Mandatory = $true)] [uint32]      $Color,
        [Parameter(Mandatory = $true)] [scriptblock] $Conf,
        [Parameter(Mandatory = $true)] [scriptblock] $Exec,
        [Parameter(Mandatory = $true)] [int]         $Interval
    )
    begin {}
    process {
        $mname = "$($Name)Launcher@$Interval)"
        $mutex = New-Object System.Threading.Mutex($false, $mname)
        try {
            # 多重起動回避
            if ($mutex.WaitOne(0, $false)) {
                try {
                    # コンテキスト作成
                    $AppCtxt = New-Object System.Windows.Forms.ApplicationContext
    
                    # タスクトレイアイコン作成
                    $TrayIcon = [System.Windows.Forms.NotifyIcon]@{
                        Icon            = GenTaskTrayIcon($Color)
                        Text            = $Name
                    }
                    $TrayIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

                    # 設定メニュー
                    $ConfMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '設定' }
                    $ConfMenu.add_Click({
                        try {
                            $null = $Conf.Invoke()
                        } catch {
                            $TrayIcon.BalloonTipText = $_.ToString()
                            $TrayIcon.ShowBalloonTip(5000)
                        }
                    })
                    $TrayIcon.ContextMenuStrip.Items.Add($ConfMenu) > $null

                    # 実行メニュー
                    $ExecMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '実行' }
                    $ExecMenu.add_Click({
                        $rsl = ""
                        try {
                            $rsl = $Exec.Invoke()
                        } catch {
                            $TrayIcon.BalloonTipText = $_.ToString()
                            $TrayIcon.ShowBalloonTip(5000)
                        }
                        if ($rsl -ne "") {
                            $TrayIcon.BalloonTipText = $rsl
                            $TrayIcon.ShowBalloonTip(5000)
                        }
                    })
                    $TrayIcon.ContextMenuStrip.Items.Add($ExecMenu) > $null

                    # 終了メニュー
                    $ExitMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '終了' }
                    $ExitMenu.add_Click({
                        $AppCtxt.ExitThread()
                    })
                    $TrayIcon.ContextMenuStrip.Items.Add($ExitMenu) > $null
                    
                    # インターバル
                    if ($Interval -gt 0){
                        $TrayTimer = New-Object Windows.Forms.Timer
                        $TrayTimer.Add_Tick({
                            $TrayTimer.Stop()
                            $rsl = ""
                            try {
                                $rsl = $Exec.Invoke()
                            } catch {
                                $TrayIcon.BalloonTipText = $_.ToString()
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                            if ($rsl -ne "") {
                                $TrayIcon.BalloonTipText = $rsl
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                            $TrayTimer.Interval = $Interval
                            $TrayTimer.Start()
                        })
                        $TrayTimer.Interval = 5000 # 固定
                        $TrayTimer.Enabled = $true
                        $TrayTimer.Start()
                    }
    
                    # タスクトレイアイコン登録
                    $TrayIcon.Visible = $true
                    [System.Windows.Forms.Application]::Run($AppCtxt) > $null
                    $TrayIcon.Visible = $false
                    $TrayTimer.Stop()
                } finally {
                    if($TrayTimer){$TrayTimer.Dispose()} 
                    if($TrayIcon ){$TrayIcon.Dispose()} 
                    if($mutex    ){$mutex.ReleaseMutex()}
                }
            }
        } finally {
            $mutex.Dispose()
        }
    
    }
    end {}
}

<#
.SYNOPSIS
    IP Messengerでメッセージを飛ばす
.DESCRIPTION
    IP Messengerでメッセージを飛ばす
.PARAMETER TargerIP
    送信先IPorホスト名
.PARAMETER Message
    送信メッセージ(改行は文字の``\n``)
.NOTES
    依存:winget install FastCopy.IPMsg
    ※IPMessengerは単一行がファイルパスの場合その行がリンクになる
#>
function SendIPMsg {
    param (
        [Parameter(Mandatory = $false)] [string] $ExePath = "$ENV:USERPROFILE\AppData\Local\IPMsg\IPMsg.exe",
        [Parameter(Mandatory = $false)] [string] $TargerIP = "127.0.0.1",
        [Parameter(Mandatory = $true)]  [string] $Message
    )
    begin {}
    process {
        Start-Process -FilePath $ExePath -ArgumentList "/MSGEX", $TargerIP, $Message -NoNewWindow -Wait
    }
    end {}
}

## ############################################################################
## ファイル操作関連

# ユニーク名取得
function local:GenUniqName([string] $DstPath, [string] $SrcPath){
    $sUniq = $DstPath
    $lUniq = 1
    while( (Test-Path -LiteralPath $sUniq) ) {
        if ((Get-Item $SrcPath).PSIsContainer) {
            $dname = [System.IO.Path]::GetDirectoryName($DstPath)
            $fname = [System.IO.Path]::GetFileName($DstPath)
            $ename = ""
        }else{
            $dname = [System.IO.Path]::GetDirectoryName($DstPath)
            $fname = [System.IO.Path]::GetFileNameWithoutExtension($DstPath)
            $ename = [System.IO.Path]::GetExtension($DstPath)
        }
        $sUniq = [System.IO.Path]::Combine($dname, $fname + " ($lUniq)" + $ename)
        $lUniq++
    }
    return $sUniq
}

<#
.SYNOPSIS
    ファイルを重複しないファイル名にして複製します
.DESCRIPTION
    ファイルを重複しないファイル名にして複製します
.PARAMETER SrcPath
    複製元のファイルまたはフォルダのパス
.PARAMETER DstPath
    複製先のファイルまたはフォルダのパス
.EXAMPLE
    MoveItemWithUniqName -SrcPath "C:\Temp\test.txt" -DstPath "C:\Temp\test.txt" -isDir $false
    "C:\Temp\test.txt"を"C:\Temp\test.txt"に複製します
    "C:\Temp\test.txt"が既に存在する場合は"C:\Temp\test (1).txt"に複製します
#>
function CopyItemWithUniqName {
    param (
        [Parameter(Mandatory = $true)] [string] $SrcPath,
        [Parameter(Mandatory = $true)] [string] $DstPath
    )
    begin {}
    process {
        if ($SrcPath -ne $DstPath) {
            # ユニーク名取得
            $sUniq = GenUniqName $DstPath $SrcPath
            # 進捗表示
            if ((Get-Item $SrcPath).PSIsContainer) {
                $index = 0
                $count = 1
            } else {
                $index = 0
                $count = (Get-ChildItem $SrcPath -Recurse).Length
            }
            Copy-Item -LiteralPath $SrcPath -Destination $sUniq -PassThru -Recurse | 
            ForEach-Object {
                Write-Progress "$fname" -PercentComplete (($index / $count)*100)
                if ($index -lt $count){
                    $index += 1
                }
            } | Out-Null
        }
    }
    end {}
}

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
            # ユニーク名取得
            $sUniq = GenUniqName $DstPath $SrcPath
            # 進捗表示
            if ((Get-Item $SrcPath).PSIsContainer) {
                $index = 0
                $count = 1
            } else {
                $index = 0
                $count = (Get-ChildItem $SrcPath -Recurse).Length
            }
            Move-Item -LiteralPath $SrcPath -Destination $sUniq -PassThru | 
            ForEach-Object {
                Write-Progress "$fname" -PercentComplete (($index / $count)*100)
                if ($index -lt $count){
                    $index += 1
                }
            } | Out-Null
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
## Zip(Windows標準機能)

function local:innerExpand([string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath $SrcPath
    # 展開
    $null = Expand-Archive -LiteralPath """$SrcPath""" -DestinationPath """$sUniq"""
}
function local:innerCompress([string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath $SrcPath
    # 圧縮
    $null = Compress-Archive -LiteralPath """$SrcPath""" -DestinationPath """$sUniq"""
}

<#
.SYNOPSIS
    指定したパスにあるアーカイブファイルを展開します
.DESCRIPTION
    指定したパスにあるアーカイブファイルを再帰的に展開し元のアーカイブファイルを削除します
    解凍先が既にある場合上書きを自動で避けます
.PARAMETER DstPath
    展開先のディレクトリパスを指定します
.PARAMETER SrcPath
    展開するファイルまたはディレクトリのパスを指定します
.PARAMETER All
    徹底解凍するかどうか
.EXAMPLE
    ExtArc -SrcPath "C:\temp\archive.zip" -DstPath "C:\temp\extracted"
    ``C:\temp\archive.zip``を``C:\temp\extracted``に展開します
#>
function ExtArc {
    param (
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath,
        [Parameter(Mandatory = $false)] [bool]   $All = $false
    )
    begin {}
    process {
        innerExpand -DstPath $DstPath -SrcPath $SrcPath
        if ($All -eq $true){
            Get-ChildItem -LiteralPath $DstPath -File -Recurse |
            Where-Object { @(".zip") -contains $_.Extension } |
            ForEach-Object {
                $dname = [System.IO.Path]::GetDirectoryName($_.FullName)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                $ename = ""
                innerExpand -DstPath ([System.IO.Path]::Combine($dname, $fname + $ename)) -SrcPath ($_.FullName) 
                $null = Remove-Item -LiteralPath ($_.FullName) -Force
            } | Out-Null
        }
        $null = Remove-Item -LiteralPath $SrcPath -Force
    }
    end {}
}

<#
.SYNOPSIS
    指定したパスをZIP形式で圧縮します
.DESCRIPTION
    指定したパスをZIP形式で圧縮します
    圧縮先が既にある場合上書きを自動で避けます
.PARAMETER DstPath
    圧縮ファイルのパスを指定します
.PARAMETER SrcPath
    圧縮するファイルまたはディレクトリのパスを指定します
.EXAMPLE
    CmpArc -SrcPath "C:\temp\files" -DstPath "C:\temp\archive.zip"
    ``C:\temp\files``を``C:\temp\archive.zip``に圧縮します
#>
function CmpArc {
    param (
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath
    )
    begin {}
    process {
        innerCompress -DstPath $DstPath -SrcPath $SrcPath
    }
    end {}
}

## ############################################################################
## 7Zip

function local:innerExp7Z([string]$ExePath, [string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath $SrcPath
    # 展開
    $null = Start-Process -FilePath """$($ExePath)""" -WindowStyle Hidden -ArgumentList "x", """$SrcPath""", "-o""$sUniq""", "-aoa" -Wait
}
function local:innerCmp7Z([string]$ExePath, [string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath $SrcPath
    # 圧縮
    $null = Start-Process -FilePath """$($ExePath)""" -WindowStyle Hidden -ArgumentList "a", "-tzip", """$sUniq""", """$SrcPath""", "-r", "-aoa" -Wait
}

<#
.SYNOPSIS
    指定したパスにあるアーカイブファイルを展開します
.DESCRIPTION
    指定したパスにあるアーカイブファイルを再帰的に展開し元のアーカイブファイルを削除します
    解凍先が既にある場合上書きを自動で避けます
.PARAMETER ExePath
    7z.exeのパスを指定します
.PARAMETER DstPath
    展開先のディレクトリパスを指定します
.PARAMETER SrcPath
    展開するファイルまたはディレクトリのパスを指定します
.PARAMETER All
    徹底解凍するかどうか
.EXAMPLE
    ExtArc7Z -SrcPath "C:\temp\archive.zip" -DstPath "C:\temp\extracted"
    ``C:\temp\archive.zip``を``C:\temp\extracted``に展開します
.NOTES
    依存
    winget install 7zip.7zip
#>
function ExtArc7Z {
    param (
        [Parameter(Mandatory = $false)] [string]  $ExePath = "$ENV:ProgramFiles\7-Zip\7z.exe",
        [Parameter(Mandatory = $true)]  [string]  $DstPath,
        [Parameter(Mandatory = $true)]  [string]  $SrcPath,
        [Parameter(Mandatory = $false)] [bool]    $All = $false
    )
    begin {}
    process {
        innerExp7Z -ExePath $ExePath -DstPath $DstPath -SrcPath $SrcPath
        if ($All -eq $true){
            Get-ChildItem -LiteralPath $DstPath -File -Recurse |
            Where-Object { @(".zip", ".rar", ".7z") -contains $_.Extension } |
            ForEach-Object {
                $dname = [System.IO.Path]::GetDirectoryName($_.FullName)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                $ename = ""
                innerExp7Z -ExePath $ExePath -DstPath ([System.IO.Path]::Combine($dname, $fname + $ename)) -SrcPath ($_.FullName) 
                $null = Remove-Item -LiteralPath ($_.FullName) -Force
            } | Out-Null
        }
        $null = Remove-Item -LiteralPath $SrcPath -Force
    }
    end {}
}

<#
.SYNOPSIS
    指定したパスをZIP形式で圧縮します
.DESCRIPTION
    指定したパスをZIP形式で圧縮します
    圧縮先が既にある場合上書きを自動で避けます
.PARAMETER ExePath
    7z.exeのパスを指定します
.PARAMETER DstPath
    圧縮ファイルのパスを指定します
.PARAMETER SrcPath
    圧縮するファイルまたはディレクトリのパスを指定します
.EXAMPLE
    CmpArc7Z -SrcPath "C:\temp\files" -DstPath "C:\temp\archive.zip"
    ``C:\temp\files``を``C:\temp\archive.zip``に圧縮します
.NOTES
    依存
    winget install 7zip.7zip
#>
function CmpArc7Z {
    param (
        [Parameter(Mandatory = $false)] [string] $ExePath = "$ENV:ProgramFiles\7-Zip\7z.exe",
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath
    )
    begin {}
    process {
        innerCmp7Z -ExePath $ExePath -DstPath $DstPath -SrcPath $SrcPath
    }
    end {}
}

## ############################################################################
## WebAPI関連

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
