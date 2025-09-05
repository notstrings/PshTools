# Zoomミーティング作成スクリプト

# PSZoom
if (-not (Get-Module -ListAvailable -Name PSZoom)) {
    Install-Module -Name PSZoom -Scope CurrentUser -Force -AllowClobber
}
Import-Module PSZoom

# Zoom設定
$AccID = 'xxxx'
$ClientID = 'xxxx'
$ClientSecret = 'xxxx'
Connect-PSZoom -AccountID $AccID -ClientID $ClientID -ClientSecret $ClientSecret

# スケジュール取得
$Meetings = Get-ZoomMeetingsFromuser -UserId 'gijyutsu@srze.co.jp' -Type scheduled
Write-Host "=============================="
Write-Host "予定一覧:"
$Meetings.meetings |
Sort-Object start_time |
Where-Object { 
    $Time = [System.TimeZoneInfo]::ConvertTimeFromUtc($_.start_time, ([System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time")))
    $Time -gt (Get-Date)
} |
ForEach-Object {
    Write-Host ("{0:yyyy-MM-dd HH:mm}({1}分) {2}" -f $_.start_time, $_.duration, $_.topic)
}
Write-Host "=============================="

# 入力
$text = ""
$parsedDate = (Get-Date)
$duration   = 0

## タイトル
$topic = Read-Host "ミーティングのタイトルを入力してください"

## 日付時刻（YYYY-MM-DD）
do {
    $text = Read-Host "ミーティング予定日時を入力してください (例: 2025-01-01 10:00)"
    [bool]$result = [datetime]::TryParseExact(
        $text, 
        'yyyy-MM-dd HH:mm',
        [Globalization.DateTimeFormatInfo]::CurrentInfo,
        [Globalization.DateTimeStyles]::AllowWhiteSpaces,
        [ref]$parsedDate
    )
} until ($result)

## 時間（分）
do {
    $duration = 0
    $text = Read-Host "ミーティングの長さ（分）を入力してください"
    [int]::TryParse($text, [ref]$duration)
} until ($duration -ge 30)

# 確認
Write-Host "=============================="
Write-Host "以下の内容でミーティングを作成します"
Write-Host "タイトル   : $topic"
Write-Host "開始日時   : $parsedDate"
Write-Host "時間       : $duration 分"
Write-Host "=============================="
$confirm = Read-Host "実行しますか？ (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "キャンセルしました。" -ForegroundColor Yellow
    exit
}

# 実行
$meeting = New-ZoomMeeting `
    -UserId 'gijyutsu@srze.co.jp' `
    -Topic $topic `
    -StartTime $parsedDate.ToString("yyyy-MM-ddTHH:mm:ss") `
    -Duration $duration `
    -Timezone 'Asia/Tokyo'

# 結果表示
$ResultText = @"
ミーティングID: $($meeting.id)
タイトル　: $topic
開始日時　: $($meeting.start_time)
時間　　　: $($meeting.duration)分
参加URL 　: $($meeting.join_url)
パスワード: $($meeting.password)
"@
$ResultText | Set-Clipboard
Write-Host "設定完了"
Write-Host $ResultText
Write-Host "ミーティング情報をクリップボードにコピーしました！"
pause
