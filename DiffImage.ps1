$ErrorActionPreference = "Stop"

function local:Setup() {
    if ((Get-Command scoop -ErrorAction SilentlyContinue) -eq $false) {
        Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://get.scoop.sh')
    }
    scoop bucket add extras
    scoop install imagemagick
}

function local:DiffImage([System.IO.FileInfo] $LHS, [System.IO.FileInfo] $RHS) {
    try {
        $IMPath = "magick.exe"
        $Lsrcpath = $LHS.FullName
        $Rsrcpath = $RHS.FullName
        if ($LHS.LastWriteTime -le $RHS.LastWriteTime) {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$Lsrcpath"" -resize 800x800 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$Rsrcpath"" -resize 800x800 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        } else {
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$Rsrcpath"" -resize 800x800 -type GrayScale +level-colors Red,White  ""tempLHS.png"""
            $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""$Lsrcpath"" -resize 800x800 -type GrayScale +level-colors Blue,White ""tempRHS.png"""
        }
        $null = Start-Process -NoNewWindow -Wait -FilePath """$IMPath""" -ArgumentList "convert ""tempLHS.png"" ""tempRHS.png"" -compose Multiply -gravity center -composite ""diff.png"""
        $null = Start-Process -NoNewWindow -Wait -FilePath "mspaint.exe" -ArgumentList "diff.png"
    } finally {
        $null = Remove-Item -Path @("diff.png", "tempLHS.png", "tempRHS.png") -Force
    }
}

# $args = @("$($ENV:USERPROFILE)\Desktop\新しいフォルダー\aaa.png", "$($ENV:USERPROFILE)\Desktop\新しいフォルダー\bbb.png")

try {
    $null = Write-Host "---DiffImage---"
	# 引数確認
    if ($args.Length -eq 1) {
        exit
    }
	# 処理実行
    $exist = $true
    $exist = $exist -and (Test-Path -LiteralPath $args[0])
    $exist = $exist -and (Test-Path -LiteralPath $args[1])
    if ($exist) {
        DiffImage (Get-Item $args[0]) (Get-Item $args[1])
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
