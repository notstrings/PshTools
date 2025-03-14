﻿## ############################################################################
## とりあえずの関数置き場

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Drawing
Add-type -AssemblyName System.Windows.Forms

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
    クラスのメンバ変数ではなくプロパティを対象にしています
#>
function DeepCopyObj {
    param (
        [Parameter(Mandatory = $false)] [object] $Data
    )
    begin {}
    process {
        if ($null -eq $Data) {
            return $null
        } elseif ($Data.GetType().IsPrimitive) {
            return $Data
        } elseif ($Data -is [string] -or $Data -is [datetime] -or $Data -is [decimal]) {
            return $Data
        } elseif ($Data -is [hashtable]) {
            $inst = @{}
            foreach ($key in $Data.Keys) {
                $inst[$key] = DeepCopyObj $Data[$key]
            }
            return $inst
        } elseif ($Data -is [array]) {
            $inst = @()
            foreach ($item in $Data) {
                $inst += DeepCopyObj $item
            }
            return $inst
        } elseif ($Data -is [System.Management.Automation.PSCustomObject]) {
            $inst = [PSCustomObject]@{}
            foreach ($prop in $Data.PSObject.Properties) {
                $inst | Add-Member -MemberType NoteProperty -Name $prop.Name -Value (DeepCopyObj $prop.Value)
            }
            return $inst
        } elseif ($Data.GetType().IsClass -and -not $Data.GetType().IsValueType) {
            $type = $Data.GetType()
            $inst = New-Object -TypeName $type.FullName
            foreach ($prop in $type.GetProperties()){
                if ($prop.CanRead -and $prop.CanWrite) {
                    $prop.SetValue($inst, $prop.GetValue($Data))
                }
            }
            return $inst
        } else {
            return $Data
        }
    }
    end {}
}

<#
.SYNOPSIS
    PSCustomObjectを指定クラスに変換します
.DESCRIPTION
    PSCustomObjectを指定クラスに変換します
.PARAMETER Type
    変換インスタンスの型
.PARAMETER Data
    変換インスタンスのプロパティに対応するPSCustomObject
.NOTES
    クラスのメンバ変数ではなくプロパティを対象にしています
#>
function ConvertFromPSCO {
    param (
        [Parameter(Mandatory = $true)]  [System.Type]    $Type,
        [Parameter(Mandatory = $false)] [PSCustomObject] $Data
    )
    begin {}
    process {
        if ($null -eq $Data) {
            return $null
        }
        if ($Type.IsPrimitive) {
            return $Data
        } elseif ($Type.IsEnum) {
            return $Data
        } elseif (($Type.Name -eq "string") -or ($Type.Name -eq "datetime") -or ($Type.Name -eq "decimal")) {
            return $Data
        } elseif ($Type.IsArray) {
            $inst = @()
            foreach ($elm in $Data) {
                $inst += ConvertFromPSCO -Type $Type.GetElementType() -Data $elm
            }
            return $inst
        } elseif ($Type.IsClass -and -not $Type.IsValueType) {
            $inst = New-Object -TypeName $Type.FullName
            $Data.PSObject.Properties | ForEach-Object {
                $prop = $Type.GetProperty($_.Name)
                if ($null -ne $prop -and $prop.CanRead -and $prop.CanWrite) {
                    $inst.($_.Name) = ConvertFromPSCO -Type $Type.GetProperty($_.Name).PropertyType -Data $Data.($_.Name)
                }
            }
            return $inst
        } else {
            return $Data
        }
        return $null
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
        $Text = [regex]::Replace($Text, "[$Chars]+",{
            param($match)
            return [Microsoft.VisualBasic.Strings]::StrConv($match, [Microsoft.VisualBasic.VbStrConv]::Narrow)
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        return $Text
    }
    end {}
}

<#
.SYNOPSIS
    半角カナを全角に変換します
.DESCRIPTION
    半角カナを全角に変換します
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
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        return $Text
    }
    end {}
}

# 和暦(gg)が含まれているとおかしくなる対策
function local:FormatDate([datetime] $Date, [string] $Format) {
    $info = New-Object CultureInfo("ja-jp", $true)
    $info.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
    if ($Format.Contains("g")) {
        $Date.ToString($Format, $info)
    } else {
        $Date.ToString($Format)
    }
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
        [parameter(Mandatory=$false)] [string]$Format = "yyyyMMdd",
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
            if($date){ return (FormatDate $date $Format) }else{ return $match.Value }
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        ## YYYY年MM月DD日
        $Text = [regex]::Replace($Text, "(?<![0-9]+)(19|20)(\d\d)年([1-9]|0[1-9]|1[0-2])月([1-9]|0[1-9]|[12][0-9]|3[01])日",{
            param($match)
            $name = $match.Value.ToUpper()
            $date = [DateTime]::ParseExact($name, "yyyy年M月d日", $null)
            if($date){ return (FormatDate $date $Format) }else{ return $match.Value }
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
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
            if($date){ return (FormatDate $date $Format) }else{ return $match.Value }
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
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
            if($date){ return (FormatDate $date $Format) }else{ return $match.Value }
        }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
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
                    return (FormatDate $dateyy $Format)
                }elseif( ($dategg) -and ($RefDate.Year -eq $dategg.Year) ){
                    return (FormatDate $dategg $Format)
                }else{
                    return $match.Value
                }
            }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
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
                    return (FormatDate $dateyy $Format)
                }elseif( ($dategg) -and ($RefDate.Year -eq $dategg.Year) ){
                    return (FormatDate $dategg $Format)
                }else{
                    return $match.Value
                }
            }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
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
        [Parameter(Mandatory = $false)] [string] $Text
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
        [Parameter(Mandatory = $false)] [string] $Text
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
        [Parameter(Mandatory = $false)] [string] $Text
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
Function AutoGuessEncodingFileSimple {
    param (
        [Parameter(Mandatory = $true)] [string] $Path
    )
    begin {}
    process {
        return AutoGuessEncodingByteSimple ([System.IO.File]::ReadAllBytes($Path))
    }
    end {}
}

<#
.SYNOPSIS
    テキストのエンコーディングを判定します
.DESCRIPTION
    指定されたテキストのエンコーディングをBOMの有無や文字コードの出現頻度に基づいて判定します
    判定対象は日本語のテキスト(Shift-JIS/EUC-JP/UTF-8)だけですので過信しないように
.PARAMETER TextData
    エンコーディングを判定するテキストを指定します
#>
Function AutoGuessEncodingByteSimple {
    param (
        [Parameter(Mandatory = $true)] [byte[]] $TextData
    )
    begin {}
    process {
        # BOMによるエンコード判定
        foreach ($elm in [System.Text.Encoding]::GetEncodings()) {
            $oAmb = $elm.GetEncoding().GetPreamble()
            if ($oAmb.Length -gt 0 ) {
                if (Compare-Object $TextData[0..($oAmb.Length - 1)] $oAmb -IncludeEqual -ExcludeDifferent) {
                    return $elm.GetEncoding()
                }
            }
        }
        # BOM無しなので日本語前提にSJIS/EUC/UTF8の何れかをソレっぽさで判定する
        # 元ネタは特徴的な文字コードを数え上げてスコアリングしてたが...
        # ・スコアリングせずソレっぽい位置を数えるように変更
        # ・末尾のややこしい処理は移植面倒なのでバッサリ破棄
        # 元々の処理が推定だし単発スクリプトならコレで良いでしょ多分
        $SJISCount = 0..($TextData.Length - 2) | Where-Object {
            $b1 = $TextData[$_ + 0]
            $b2 = $TextData[$_ + 1]
            ((0x81 -le $b1 -and $b1 -le 0x9F) -or (0xE0 -le $b1 -and $b1 -le 0xFC)) -and
            ((0x40 -le $b2 -and $b2 -le 0x7E) -or (0x80 -le $b2 -and $b2 -le 0xFC))
        }
        $EUCJCount = 0..($TextData.Length - 2) | Where-Object {
            $b1 = $TextData[$_ + 0]
            $b2 = $TextData[$_ + 1]
            (((0xA1 -le $b1 -and $b1 -le 0xFE) -and (0xA1 -le $b2 -and $b2 -le 0xFE)) -or
             ((0x8E -eq $b1)                   -and (0xA1 -le $b2 -and $b2 -le 0xDF)))
        }
        $UTF8Count = 0..($TextData.Length - 2) | Where-Object {
            $b1 = $TextData[$_ + 0]
            $b2 = $TextData[$_ + 1]
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
    問合ダイアログを表示します
.PARAMETER Title
    ダイアログのタイトルを指定します
.PARAMETER Message
    ダイアログのメッセージを指定します
.PARAMETER Default
    デフォルトボタンを指定します
#>
function AskBox {
    param (
        [Parameter(Mandatory = $true)]  [string] $Title,
        [Parameter(Mandatory = $true)]  [string] $Message,
        [Parameter(Mandatory = $false)] [string] $Default = "Button1"
    )
    begin {}
    process {
        try {
            $DUMY = New-Object Windows.Forms.Form
            $DUMY.TopMost = $true
            $ret = [System.Windows.Forms.MessageBox]::Show( `
                $DUMY,
                $Message, `
                $Title, `
                [System.Windows.Forms.MessageBoxButtons]::YesNo, `
                [System.Windows.Forms.MessageBoxIcon]::Question, `
                $Default `
            )
            return $ret
        } finally {
            if ($null -ne $DUMY) {$DUMY.Dispose()}
        }
    }
    end {}
}

<#
.SYNOPSIS
    情報ダイアログを表示します
.PARAMETER Title
    ダイアログのタイトルを指定します
.PARAMETER Message
    ダイアログのメッセージを指定します
.PARAMETER Default
    デフォルトボタンを指定します
#>
function InfBox {
    param (
        [Parameter(Mandatory = $true)]  [string] $Title,
        [Parameter(Mandatory = $true)]  [string] $Message,
        [Parameter(Mandatory = $false)] [string] $Default = "Button1"
    )
    begin {}
    process {
        try {
            $DUMY = New-Object Windows.Forms.Form
            $DUMY.TopMost = $true
            $ret = [System.Windows.Forms.MessageBox]::Show( `
                $DUMY,
                $Message, `
                $Title, `
                [System.Windows.Forms.MessageBoxButtons]::OK, `
                [System.Windows.Forms.MessageBoxIcon]::Information, `
                $Default `
            )
            return $ret
        } finally {
            if ($null -ne $DUMY) {$DUMY.Dispose()}
        }
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
        try {
            $DUMY = New-Object Windows.Forms.Form
            $DUMY.TopMost = $true
            $FDlg = New-Object System.Windows.Forms.OpenFileDialog
            $FDlg.Title            = $Title
            $FDlg.InitialDirectory = $InitialDirectory
            $FDlg.Filter           = $Filter
            $FDlg.Multiselect      = $Multiselect
            $null = $FDlg.ShowDialog($DUMY)
            return $FDlg.FileNames
        } finally {
            if ($null -ne $DUMY) {$DUMY.Dispose()}
        }
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
        try {
            $DUMY = New-Object Windows.Forms.Form
            $DUMY.TopMost = $true
            $FDlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $FDlg.Description      = $Description
            $FDlg.InitialDirectory = $InitialDirectory
            $null = $FDlg.ShowDialog($DUMY)
            return $FDlg.SelectedPath
        } finally {
            if ($null -ne $DUMY) {$DUMY.Dispose()}
        }
    }
    end {}
}

# FileListDialog用ファイルリスト生成
function local:AddFileList([System.Windows.Forms.ListBox] $ListBox, [string[]] $FilePaths, [string] $FileFilter) {
    foreach ($FilePath in $FilePaths) {
        if (Test-Path -LiteralPath $FilePath) {
            if ([System.IO.Directory]::Exists($FilePath)) {
                $ChildFilePaths = @()
                @(Get-ChildItem -LiteralPath $FilePath -File -Recurse) | ForEach-Object {
                    $ChildFilePaths += $_.FullName
                }
                AddFileList $ListBox $ChildFilePaths $FileFilter
            } else {
                if ([System.IO.Path]::GetFileName($FilePath) -match $FileFilter) {
                    if ($ListBox.Items -notcontains $FilePath) {
                        [void]$ListBox.Items.Add($FilePath)
                    }
                }
            }
        }
    }
}

<#
.SYNOPSIS
    ドラッグ＆ドロップで受け取ったファイルを選択するためのダイアログを表示します
.DESCRIPTION
    タイトル/メッセージ/ファイルフィルター/初期リストを指定して
    ドラッグ＆ドロップで受け取ったファイルを表示するダイアログボックスを作成します
    ユーザーがOKボタンを押すと選択した結果/ファイルリスト/選択中ファイル名が返されます
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
        $frmMain = New-Object System.Windows.Forms.Form -Property @{
            Text          = $Title                                                      # タイトル
            StartPosition = 'CenterScreen'                                              # 表示位置
            Size          = New-Object System.Drawing.Size(480,320)
            Padding       = New-Object System.Windows.Forms.Padding(5)
        }

        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
            Dock     = [System.Windows.Forms.DockStyle]::Fill
            RowCount = 2
        }

        $pnlBody = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $lblDD = New-Object System.Windows.Forms.Label -Property @{
            Dock      = [System.Windows.Forms.DockStyle]::Top
            Text      = $Message
            AutoSize  = $true
            TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        }
        $lbxDD = New-Object System.Windows.Forms.ListBox -Property @{
            Dock        = [System.Windows.Forms.DockStyle]::Fill
            AllowDrop   = $true
        }

        $pnlTail = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $btnOK = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "OK"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::OK
        }
        $btnCancel = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "Cancel"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::Cancel
        }

        $null = $pnlBody.Controls.Add($lbxDD)
        $null = $pnlBody.Controls.Add($lblDD)
        $null = $pnlTail.Controls.Add($btnOK)
        $null = $pnlTail.Controls.Add($btnCancel)
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)

        $null = $frmMain.Add_Load({
            $frmMain.BringToFront()
        })
        $null = $lbxDD.Add_DragEnter({
            $_.Effect = "All"
        })
        $null = $lbxDD.Add_DragDrop({
            AddFileList $lbxDD $_.Data.GetData("FileDrop") $FileFilter
        })
        $null = $lbxDD.Add_KeyDown({
            if ($_.KeyCode -eq "Delete") {
                if ($lbxDD.SelectedIndex -ge 0){
                    [void]$lbxDD.Items.RemoveAt($lbxDD.SelectedIndex)
                }
            }
        })

        # フォーム表示
        if ($null -ne $FileList) {
            AddFileList $lbxDD $FileList $FileFilter
        }
        $frmMain.AcceptButton = $btnOK
        $frmMain.CancelButton = $btnCancel
        $null = $frmMain.ShowDialog()
        $item = ""
        if ($lbxDD.SelectedIndex -ge 0){
            $item = $lbxDD.Items[$lbxDD.SelectedIndex]
        }
        return @($frmMain.DialogResult, $lbxDD.Items, $item)
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
    ユーザーがOKボタンを押すと選択した結果/ファイルリスト/選択中ファイル名が返されます
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
        $frmMain = New-Object System.Windows.Forms.Form -Property @{
            Text          = $Title                                                      # タイトル
            StartPosition = 'CenterScreen'                                              # 表示位置
            Size          = New-Object System.Drawing.Size(480,320)
            Padding       = New-Object System.Windows.Forms.Padding(5)
        }

        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
            Dock     = [System.Windows.Forms.DockStyle]::Fill
            RowCount = 2
        }

        $pnlBody = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $lblDD = New-Object System.Windows.Forms.Label -Property @{
            Dock      = [System.Windows.Forms.DockStyle]::Top
            Text      = $Message
            AutoSize  = $true
            TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        }
        $lbxDD = New-Object System.Windows.Forms.ListBox -Property @{
            Dock        = [System.Windows.Forms.DockStyle]::Fill
            AllowDrop   = $true
        }
        $grpOpt = New-Object System.Windows.Forms.GroupBox -Property @{
            Dock     = [System.Windows.Forms.DockStyle]::Right
            Text     = "Options"
            AutoSize = $true
            Padding  = New-Object System.Windows.Forms.Padding(5)
        }
        $flpOpt = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{
            Dock          = [System.Windows.Forms.DockStyle]::Fill
            AutoSize      = $true
            FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        }
        $Checked = $true
        $Options | ForEach-Object {
            $rdoOpt = New-Object System.Windows.Forms.RadioButton -Property @{
                Text     = $_
                Checked  = $Checked
                AutoSize = $true
            }
            $flpOpt.Controls.Add($rdoOpt)
            $Checked = $false
        }

        $pnlTail = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $btnOK = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "OK"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::OK
        }
        $btnCancel = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "Cancel"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::Cancel
        }

        $null = $pnlBody.Controls.Add($lbxDD)
        $null = $pnlBody.Controls.Add($lblDD)
        $null = $pnlBody.Controls.Add($grpOpt)
        $null = $grpOpt.Controls.Add($flpOpt)
        $null = $pnlTail.Controls.Add($btnOK)
        $null = $pnlTail.Controls.Add($btnCancel)
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)

        $null = $frmMain.Add_Load({
            $frmMain.BringToFront()
        })
        $null = $lbxDD.Add_DragEnter({
            $_.Effect = "All"
        })
        $null = $lbxDD.Add_DragDrop({
            AddFileList $lbxDD $_.Data.GetData("FileDrop") $FileFilter
        })
        $null = $lbxDD.Add_KeyDown({
            if ($_.KeyCode -eq "Delete") {
                if ($lbxDD.SelectedIndex -ge 0){
                    [void]$lbxDD.Items.RemoveAt($lbxDD.SelectedIndex)
                }
            }
        })

        # フォーム表示
        if ($null -ne $FileList) {
            $null = AddFileList $lbxDD $FileList $FileFilter
        }
        $frmMain.AcceptButton = $btnOK
        $frmMain.CancelButton = $btnCancel
        $null = $frmMain.ShowDialog()
        $item = ""
        if ($lbxDD.SelectedIndex -ge 0){
            $item = $lbxDD.Items[$lbxDD.SelectedIndex]
        }
        return @($frmMain.DialogResult, $lbxDD.Items, $item, ($flpOpt.Controls | Where-Object {$_.Checked -eq $true} | Select-Object -ExpandProperty Text))
    }
    end {}
}

<#
.SYNOPSIS
    プロパディグリッドを使った汎用設定ダイアログを表示します
.DESCRIPTION
    プロパディグリッドを使った汎用設定ダイアログを表示します
    設定対象が参照型の場合は画面操作による副作用を受けますので編集用変数を用意してください
.PARAMETER Title
    ダイアログボックスのタイトルに設定する文字列です
.PARAMETER Setting
    設定対象オブジェクト(クラスインスタンスを想定)
.EXAMPLE
    Add-Type -AssemblyName "System.ComponentModel"          # 
    Add-Type -AssemblyName "System.Drawing"                 # 
    Add-Type -AssemblyName "System.Windows.Forms.Design"    # PowerShell5では使えないorz
    Invoke-Expression -Command @"
    class AppSettings {
        [System.ComponentModel.Description("名前")]
        [string]$AppName
        [int]$Version
        [bool]$AutoUpdate

        # ファイル/フォルダパスの場合※PowerShell7以降用
        # [System.ComponentModel.Editor(([System.Windows.Forms.Design.FileNameEditor]), ([System.Drawing.Design.UITypeEditor]))]    # ファイルの場合
        # [System.ComponentModel.Editor(([System.Windows.Forms.Design.FolderNameEditor]), ([System.Drawing.Design.UITypeEditor]))]  # フォルダの場合
        [string]$LogFilePath

        # Enumはコンボになる
        [System.Diagnostics.SourceLevels]$LogLevel

        # 子要素に直接クラスがある場合は展開用の属性が必要だが子要素がクラスの配列の場合はこんなことしなくていい
        # [System.ComponentModel.TypeConverter(([System.ComponentModel.ExpandableObjectConverter]))]
        # [ChildNode] elm
    }
    "@
    $settings = [AppSettings]@{
        AppName = "My Application"
        Version = 1
        AutoUpdate = $true
        LogFilePath = "C:\app.log"
        LogLevel = [System.Diagnostics.SourceLevels]::Information
    }
    $edit = ConvertFromPSCO ([AutoRenameConf]) $settings
    $ret = ShowSettingDialog "Title" $edit
    if ($ret -eq "OK") {
        $edit
    }
#>
function ShowSettingDialog {
    param (
        [Parameter(Mandatory = $true)] [string]        $Title,
        [Parameter(Mandatory = $true)] [System.Object] $Setting
    )
    begin {}
    process {
        # フォーム生成
        $frmMain = New-Object System.Windows.Forms.Form -Property @{
            Text          = $Title                                                      # タイトル
            StartPosition = 'CenterScreen'                                              # 表示位置
            Size          = New-Object System.Drawing.Size(480,320)
            Padding       = New-Object System.Windows.Forms.Padding(5)
        }

        $tlpMain = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
            Dock     = [System.Windows.Forms.DockStyle]::Fill
            RowCount = 2
        }

        $pnlBody = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $grdProp = New-Object System.Windows.Forms.PropertyGrid -Property @{
            Dock           = [System.Windows.Forms.DockStyle]::Fill
            SelectedObject = $Setting
        }

        $pnlTail = New-Object System.Windows.Forms.Panel -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Fill
        }
        $btnOK = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "OK"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::OK
        }
        $btnCancel = New-Object System.Windows.Forms.Button -Property @{
            Dock                    = [System.Windows.Forms.DockStyle]::Right
            Size                    = New-Object System.Drawing.Size(128, 0) # ボタン巾のみ指定可能
            Text                    = "Cancel"
            UseVisualStyleBackColor = $true
            DialogResult            = [Windows.Forms.DialogResult]::Cancel
        }

        $null = $pnlBody.Controls.Add($grdProp)
        $null = $pnlTail.Controls.Add($btnOK)
        $null = $pnlTail.Controls.Add($btnCancel)
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        $null = $tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) # ボタン高さはコレ
        $null = $tlpMain.Controls.Add($pnlBody, 0, 0)
        $null = $tlpMain.Controls.Add($pnlTail, 0, 1)
        $null = $frmMain.Controls.Add($tlpMain)

        $null = $frmMain.Add_Load({
            $frmMain.BringToFront()
        })

        # フォーム表示
        $null = $frmMain.ShowDialog()

        return $frmMain.DialogResult
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
.PARAMETER MenuNameExit
    コンテキスメニュー表示文字列(Exit)
.PARAMETER MenuNameConf
    コンテキスメニュー表示文字列(Conf)
.PARAMETER MenuNameExec
    コンテキスメニュー表示文字列(Exec)
.NOTES
    元ネタ:https://aquasoftware.net/blog/?p=1244
#>
function RunInTaskTray {
    param (
        [Parameter(Mandatory = $true)]  [string]      $Name,
        [Parameter(Mandatory = $true)]  [uint32]      $Color,
        [Parameter(Mandatory = $true)]  [scriptblock] $Conf,
        [Parameter(Mandatory = $true)]  [scriptblock] $Exec,
        [Parameter(Mandatory = $true)]  [int]         $Interval,
        [Parameter(Mandatory = $false)] [string]      $MenuNameExit = "終了",
        [Parameter(Mandatory = $false)] [string]      $MenuNameConf = "設定",
        [Parameter(Mandatory = $false)] [string]      $MenuNameExec = "実行"
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
                        BalloonTipIcon  = 'Error'
                        BalloonTipTitle = 'Error'
                    }
                    $TrayIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

                    # 設定メニュー
                    if ($MenuNameConf) {
                        $ConfMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = $MenuNameConf }
                        $ConfMenu.add_Click({
                            try {
                                $null = $Conf.Invoke()
                            } catch {
                                $TrayIcon.BalloonTipIcon = "Error"
                                $TrayIcon.BalloonTipText = $_.ToString()
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                        })
                        $TrayIcon.ContextMenuStrip.Items.Add($ConfMenu) > $null
                    }

                    # 実行メニュー
                    if ($MenuNameExec) {
                        $ExecMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = $MenuNameExec }
                        $ExecMenu.add_Click({
                            $rsl = ""
                            try {
                                $rsl = $Exec.Invoke()
                            } catch {
                                $TrayIcon.BalloonTipIcon = "Error"
                                $TrayIcon.BalloonTipText = $_.ToString()
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                            if ($rsl -ne "") {
                                $TrayIcon.BalloonTipIcon = "Info"
                                $TrayIcon.BalloonTipText = $rsl
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                        })
                        $TrayIcon.ContextMenuStrip.Items.Add($ExecMenu) > $null
                    }

                    # 終了メニュー
                    if ("" -eq $ExitMenu -or $null -eq $ExitMenu) {
                        $ExitMenu = "Exit"
                    }
                    $ExitMenu = [System.Windows.Forms.ToolStripMenuItem]@{ Text = $MenuNameExit }
                    $ExitMenu.add_Click({
                        $AppCtxt.ExitThread()
                    })
                    $TrayIcon.ContextMenuStrip.Items.Add($ExitMenu) > $null

                    # インターバル
                    $TrayTimer = New-Object Windows.Forms.Timer
                    if ($Interval -gt 0){
                        $TrayTimer.Add_Tick({
                            $TrayTimer.Stop()
                            $rsl = ""
                            try {
                                $rsl = $Exec.Invoke()
                            } catch {
                                $TrayIcon.BalloonTipIcon = "Error"
                                $TrayIcon.BalloonTipText = $_.ToString()
                                $TrayIcon.ShowBalloonTip(5000)
                            }
                            if ($rsl -ne "") {
                                $TrayIcon.BalloonTipIcon = "Info"
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
    トーストを表示する
.DESCRIPTION
    トーストを表示する
    PowerShell5/7系どっちでも特に問題なく表示できる
.PARAMETER Title
    タイトル
.PARAMETER Message
    メッセージ
.PARAMETER Detail
    詳細
.EXAMPLE
    ShowToast "aaa" "bbb" "ccc"
#>
function ShowToast {
    param (
        [Parameter(Mandatory = $true)]  [string] $Title,
        [Parameter(Mandatory = $true)]  [string] $Message,
        [Parameter(Mandatory = $false)] [string] $Detail
    )
    begin {}
    process {
        $null = powershell {
            param ($Title, $Message, $Detail)
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
            $app_id = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
            $content = ""
            $content += "<?xml version='1.0' encoding='utf-8'?>"
            $content += "<toast>"
            $content += "  <visual>"
            $content += "    <binding template='ToastGeneric'>"
            $content += "        <text>$Title</text>"
            $content += "        <text>$Message</text>"
            $content += "        <text>$Detail</text>"
            $content += "    </binding>"
            $content += "  </visual>"
            $content += "</toast>"
            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($content)
            $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app_id).Show($toast)
        } -args $Title, $Message, $Detail
    }
    end {}
}

<#
.SYNOPSIS
    Mailを送信します
.DESCRIPTION
    指定されたMailアカウントを使用してメールを送信します
    送信元/送信先/CC/BCC/件名/本文/添付ファイルを指定してメールを送信できます
.PARAMETER Subject
    メールの件名です
.PARAMETER Body
    メールの本文です
.PARAMETER From
    送信元のメールアドレスです
.PARAMETER To
    送信先のメールアドレスです
    ※CC/BCCと異なりTOは宗教上の理由で単一指定です
.PARAMETER CC
    CCに追加するメールアドレスです
.PARAMETER BCC
    BCCに追加するメールアドレスです
.PARAMETER Attachment
    添付ファイルのパスです
.PARAMETER SMTPServer
    SMTPサーバーのホスト名です
.PARAMETER SMTPPort
    SMTPサーバーのポート番号です
.PARAMETER SMTPSSL
    SMTPサーバーのSSL使用可否です
.PARAMETER SMTPUID
    GmailアカウントのユーザーIDです
.PARAMETER SMTPPWD
    Gmailアカウントのパスワードです
#>
function SendMail {
    param (
        [Parameter(Mandatory = $true)]  [string]   $Subject,
        [Parameter(Mandatory = $true)]  [string]   $Body,
        [Parameter(Mandatory = $true)]  [string]   $From,
        [Parameter(Mandatory = $true)]  [string]   $To,
        [Parameter(Mandatory = $false)] [string[]] $CC = @(),
        [Parameter(Mandatory = $false)] [string[]] $BCC = @(),
        [Parameter(Mandatory = $false)] [string[]] $Attachment = @(),
        [Parameter(Mandatory = $false)] [string]   $SMTPServer = "smtp.gmail.com",
        [Parameter(Mandatory = $false)] [int]      $SMTPPort = 587,
        [Parameter(Mandatory = $false)] [boolean]  $SMTPSSL = $true,
        [Parameter(Mandatory = $true)]  [string]   $SMTPUID,
        [Parameter(Mandatory = $true)]  [string]   $SMTPPWD
    )
    begin {}
    process {
        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SMTPPort)
        $SMTPClient.EnableSsl = $SMTPSSL
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUID, $SMTPPWD)
        $Mail = New-Object Net.Mail.MailMessage($From,$To,$Subject,$Body)
        $CC | ForEach-Object {
            $Mail.CC.Add($_)
        }
        $BCC | ForEach-Object {
            $Mail.BCC.Add($_)
        }
        $Attachment | ForEach-Object {
            $Mail.Attachments.Add((New-Object Net.Mail.Attachment($_)))
        }
        $SMTPClient.Send($Mail)
    }
    end {}
}

<#
.SYNOPSIS
    IP Messengerでメッセージを飛ばす
.DESCRIPTION
    IP Messengerでメッセージを飛ばす
.PARAMETER ExePath
    IPMsg実行ファイルのある場所
.PARAMETER TargerIP
    送信先IPorホスト名
.PARAMETER Message
    送信メッセージ(改行は文字の``\n``)
.NOTES
    依存:winget install FastCopy.IPMsg
#>
function SendIPMsg {
    param (
        [Parameter(Mandatory = $false)] [string] $ExePath = "$ENV:USERPROFILE\AppData\Local\IPMsg\IPMsg.exe",
        [Parameter(Mandatory = $false)] [string] $TargerIP = "127.0.0.1",
        [Parameter(Mandatory = $true)]  [string] $Message
    )
    begin {}
    process {
        $Message = $Message.Replace("`r`n","\n")
        $Message = $Message.Replace("`n"  ,"\n")
        Start-Process -NoNewWindow -FilePath $ExePath -ArgumentList "/MSGEX", $TargerIP, $Message -Wait
    }
    end {}
}

<#
.SYNOPSIS
    IP Messengerでメッセージを飛ばす
.DESCRIPTION
    IP Messengerでメッセージを飛ばす
.PARAMETER HostName
    送信先ホスト名
.PARAMETER UserName
    ユーザ名
.PARAMETER GroupName
    グループ名
.PARAMETER Message
    送信メッセージ(改行はそのまま``\n``)
.NOTES
    バイナリ依存なしだが送信ログとかは一切無い
#>
function SendRawIPMsg {
    param (
        [Parameter(Mandatory = $false)] [string] $HostName = "localhost",
        [Parameter(Mandatory = $false)] [string] $UserName = "Notifier",
        [Parameter(Mandatory = $false)] [string] $GroupName = "PowerShell",
        [Parameter(Mandatory = $true)]  [string] $Message
    )
    begin {}
    process {
        $HostAddr = [System.Net.Dns]::GetHostAddresses($HostName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
        $EpochTim = [int][double]::Parse((Get-Date -UFormat %s))
        $text = "1:" + $EpochTim + ":" + $UserName + ":" + $GroupName + ":32:" + $Message
        $data = [System.Text.Encoding]::Convert(
            [System.Text.Encoding]::UTF8,
            [System.Text.Encoding]::GetEncoding("shift_jis"),
            [System.Text.Encoding]::UTF8.GetBytes($text)
        )
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse($HostAddr), 2425)
        $client = New-Object System.Net.Sockets.UdpClient
        $client.Send($data, $data.Length, $ep)
        $client.Close()
    }
    end {}
}

## ############################################################################
## ファイル操作関連

# ユニーク名取得
function GenUniqName {
    param (
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [bool]   $isDir
    )
    begin {}
    process {
        $sUniq = $Path
        $lUniq = 1
        while( (Test-Path -LiteralPath $sUniq) ) {
            if ($isDir) {
                $dname = [System.IO.Path]::GetDirectoryName($Path)
                $fname = [System.IO.Path]::GetFileName($Path)
                $ename = ""
            } else {
                $dname = [System.IO.Path]::GetDirectoryName($Path)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($Path)
                $ename = [System.IO.Path]::GetExtension($Path)
            }
            $sUniq = [System.IO.Path]::Combine($dname, $fname + " ($lUniq)" + $ename)
            $lUniq++
        }
        return $sUniq
    }
    end {}
}

<#
.SYNOPSIS
    ファイル・フォルダを進捗表示付きで複製します
.PARAMETER SrcPath
    複製元のファイルまたはフォルダのパス
.PARAMETER DstPath
    複製先のファイルまたはフォルダのパス
#>
function CopyItemWithProgress {
    param (
        [Parameter(Mandatory = $true)] [string] $SrcPath,
        [Parameter(Mandatory = $true)] [string] $DstPath
    )
    begin {}
    process {
        if ($SrcPath -ne $DstPath) {
            $index = 0
            $count = @(Get-ChildItem $SrcPath -Recurse).Count
            Copy-Item -LiteralPath $SrcPath -Destination $DstPath -PassThru -Recurse |
            ForEach-Object {
                Write-Progress "$fname" -PercentComplete (($index / $count)*100)
                if ($index -lt $count){ $index += 1 }
            } | Out-Null
            Write-Progress "$fname" -Completed
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
function MoveItemWithProgress {
    param (
        [Parameter(Mandatory = $true)] [string] $SrcPath,
        [Parameter(Mandatory = $true)] [string] $DstPath
    )
    begin {}
    process {
        if ($SrcPath -ne $DstPath) {
            $index = 0
            $count = @(Get-ChildItem $SrcPath -Recurse).Count
            Move-Item -LiteralPath $SrcPath -Destination $DstPath -PassThru |
            ForEach-Object {
                if ($count -gt 0) {
                    Write-Progress "$fname" -PercentComplete (($index / $count)*100)
                    if ($index -lt $count){ $index += 1 }
                }
            } | Out-Null
            Write-Progress "$fname" -Completed
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
## Zip
## ・Windows環境自体のセットアップ用
## ・Expand-ArchiveはマルチバイトがゴミなのでWindows標準機能のtarでどうにかする

function local:innerExpand([string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath ([System.IO.Directory]::Exists($DstPath))
    # 展開
    $null = New-Item $sUniq -ItemType Directory -ErrorAction SilentlyContinue
    $null = Start-Process -NoNewWindow -Wait -FilePath tar.exe -ArgumentList "-zvxf ""$SrcPath"" -C ""$sUniq"""
    return $sUniq
}
function local:innerCompress([string]$DstPath, [string]$SrcPath) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath ([System.IO.Directory]::Exists($DstPath))
    $sSrcD = [System.IO.Path]::GetDirectoryName($SrcPath)
    $sSrcF = [System.IO.Path]::GetFileName($SrcPath)
    # 圧縮
    $null = Start-Process -NoNewWindow -Wait -FilePath tar.exe -ArgumentList "-caf ""$sUniq"" -C ""$sSrcD"" ""$sSrcF"""
    return $sUniq
}

<#
.SYNOPSIS
    指定したパスにあるアーカイブファイルを展開します
.DESCRIPTION
    指定したパスにあるアーカイブファイルを再帰的に展開し元のアーカイブファイルを削除します
    展開先が既にある場合上書きを自動で避けます
.PARAMETER DstPath
    展開先のディレクトリパスを指定します
.PARAMETER SrcPath
    展開するファイルまたはディレクトリのパスを指定します
.PARAMETER All
    徹底展開するかどうか
.EXAMPLE
    ExtArc -SrcPath "C:\temp\archive.zip" -DstPath "C:\temp\extracted"
    ``C:\temp\archive.zip``を``C:\temp\extracted``に展開します
.NOTES
    tarによる圧縮展開は日本語が化ける可能性がある
#>
function ExtArc {
    param (
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath,
        [Parameter(Mandatory = $false)] [bool]   $All = $false
    )
    begin {}
    process {
        $ret = innerExpand -DstPath $DstPath -SrcPath $SrcPath
        if ($All -eq $true){
            Get-ChildItem -LiteralPath $DstPath -File -Recurse |
            Where-Object { @(".zip") -contains $_.Extension } |
            ForEach-Object {
                $dname = [System.IO.Path]::GetDirectoryName($_.FullName)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                $ename = ""
                $null = ExtArc -DstPath ([System.IO.Path]::Combine($dname, $fname + $ename)) -SrcPath ($_.FullName) -All $All
            } | Out-Null
        }
        $null = Remove-Item -LiteralPath $SrcPath -Force
        return $ret
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
.NOTES
    tarによる圧縮展開は日本語が化ける可能性がある
#>
function CmpArc {
    param (
        [Parameter(Mandatory = $true)] [string] $DstPath,
        [Parameter(Mandatory = $true)] [string] $SrcPath
    )
    begin {}
    process {
        innerCompress -DstPath $DstPath -SrcPath $SrcPath
    }
    end {}
}

## ############################################################################
## 7Zip

function local:innerExp7Z([string]$ExePath, [string]$DstPath, [string]$SrcPath, [string] $ZipPwd, [string] $FileNameEncode) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath ([System.IO.Directory]::Exists($DstPath))
    # 展開
    ## -aoa    :展開先に同名ファイルがある場合上書き
    ## -spe    :抽出コマンドのルートフォルダーの重複を除去
    ## -mcp=XXX:無指定の場合はUTFか自動判定しそれ以外はそのまんま扱う7zのデフォ動作/アホなアーカイブでSJISファイル名を強制する必要があるかも
    $arg = " x ""$SrcPath"" -o""$sUniq"" -aoa -spe "
    if ($ZipPwd -ne "") {
        $arg += " -p$ZipPwd"
    }
    if ($FileNameEncode -ne "") {
        $arg += " -mcp=$FileNameEncode"
    }
    $null = Start-Process -NoNewWindow -FilePath """$($ExePath)""" -ArgumentList $arg -Wait
    return $sUniq
}
function local:innerCmp7Z([string]$ExePath, [string]$DstPath, [string]$SrcPath, [string] $ZipPwd, [int] $DivideSize, [bool] $DelSrc) {
    # ユニーク名取得
    $sUniq = GenUniqName $DstPath ([System.IO.Directory]::Exists($DstPath))
    # 圧縮
    ## -aoa   :圧縮先に同名ファイルがある場合上書き
    ## -r0    :指定ディレクトリとサブディレクトリのみ再帰処理 ※-rは兄弟ディレクトリも含む...初見で分かるわけねぇだろ、ソレ
    ## -sdel  :圧縮後にファイルを削除
    ## -mcu=on:UTF8ファイル名で圧縮する/事実上必須オプション...ならデフォにしてくれればいいものを...
    $arg = " a -tzip ""$sUniq"" ""$SrcPath"" -aoa -r0 -mcu=on"
    if ($ZipPwd -ne "" ) {
        $arg += " -p$ZipPwd"
    }
    if ($DivideSize -gt 0 ) {
        $arg += " -v$($DivideSize)m"
    }
    if ($DelSrc -eq $true) {
        $arg += " -sdel"
    }
    $null = Start-Process -NoNewWindow -FilePath """$($ExePath)""" -ArgumentList $arg -Wait
    return $sUniq
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
    scoop install 7zip
#>
function ExtArc7Z {
    param (
        [Parameter(Mandatory = $false)] [string] $ExePath = "7z.exe",
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath,
        [Parameter(Mandatory = $false)] [bool]   $Recursive = $false,
        [Parameter(Mandatory = $false)] [string] $ZipPwd = "",
        [Parameter(Mandatory = $false)] [bool]   $DelSrc = $true,
        [Parameter(Mandatory = $false)] [string] $FileNameEncode = ""
    )
    begin {}
    process {
        $ret = innerExp7Z -ExePath $ExePath -DstPath $DstPath -SrcPath $SrcPath -ZipPwd $ZipPwd -FileNameEncode $FileNameEncode
        if ($Recursive -eq $true){
            Get-ChildItem -LiteralPath $DstPath -File -Recurse |
            Where-Object { @(".7Z", ".GZ", ".ZIP", ".BZ2", ".TAR", ".LZH", ".LZS", ".LHA", ".GZIP", ".LZMA") -contains ($_.Extension.ToUpper()) } |
            ForEach-Object {
                $dname = [System.IO.Path]::GetDirectoryName($_.FullName)
                $fname = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                $ename = ""
                $null = ExtArc7Z -ExePath $ExePath -DstPath ([System.IO.Path]::Combine($dname, $fname + $ename)) -SrcPath ($_.FullName) -Recursive $Recursive -ZipPwd $ZipPwd -FileNameEncode $FileNameEncode
            } | Out-Null
        }
        if ($DelSrc -eq $true) {
            $null = Remove-Item -LiteralPath $SrcPath -Force
        }
        return $ret
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
    scoop install 7zip
#>
function CmpArc7Z {
    param (
        [Parameter(Mandatory = $false)] [string] $ExePath = "7z.exe",
        [Parameter(Mandatory = $true)]  [string] $DstPath,
        [Parameter(Mandatory = $true)]  [string] $SrcPath,
        [Parameter(Mandatory = $false)] [string] $ZipPwd = "",
        [Parameter(Mandatory = $false)] [int]    $DivideSize = "0",
        [Parameter(Mandatory = $false)] [bool]   $DelSrc = $true
    )
    begin {}
    process {
        $ret = innerCmp7Z -ExePath $ExePath -DstPath $DstPath -SrcPath $SrcPath -ZipPwd $ZipPwd -DivideSize $DivideSize -DelSrc $DelSrc
        return $ret
    }
    end {}
}

## ############################################################################
## WebAPI関連

<#
.SYNOPSIS
    GitHubから指定ファイルをダウンロードします
.PARAMETER RepOwner
    リポジトリオーナー名
.PARAMETER RepName
    リポジトリ名
.PARAMETER Filter
    取得ファイル名のフィルタ
.PARAMETER OutputDirectory
    取得先フォルダ
#>
function DownloadGitHubLatest {
    param (
        [Parameter(Mandatory = $true)]  [string]$RepOwner,
        [Parameter(Mandatory = $true)]  [string]$RepName,
        [Parameter(Mandatory = $false)] [string]$Filter = ".*",
        [Parameter(Mandatory = $false)] [string]$OutputDirectory = $PSScriptRoot
    )
    begin {}
    process {
        $ret = @()
        $URI = "https://api.github.com/repos/$RepOwner/$RepName/releases/latest"
        $inf = Invoke-RestMethod -Uri $URI -Headers @{ "Accept" = "application/vnd.github.v3+json" }
        $ast = $inf.assets | Where-Object { $_.name -match $Filter }
        foreach ($elm in $ast) {
            $dlsrc = $elm.browser_download_url
            $dldst = Join-Path $OutputDirectory $elm.name
            $null  = New-Item $OutputDirectory -ItemType Directory -ErrorAction SilentlyContinue
            $null  = Invoke-WebRequest -Uri $dlsrc -OutFile $dldst
            $ret  += $dldst
        }
        return $ret
    }
    end {}
}

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
