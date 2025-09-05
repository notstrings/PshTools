# =========================================
# PSZoom ミーティング作成スクリプト
# =========================================

# --- モジュール確認 ---
if (-not (Get-Module -ListAvailable -Name PSZoom)) {
    Install-Module -Name PSZoom -Scope CurrentUser -Force -AllowClobber
}
Import-Module PSZoom

# --- Zoom OAuth情報 ---
$AccID = 'xxx'
$ClientID = 'xxx'
$ClientSecret = 'xxx'

Connect-PSZoom -AccountID $AccID -ClientID $ClientID -ClientSecret $ClientSecret

# --- 今後の予定取得 ---
$Meetings = Get-ZoomMeetingsFromuser -UserId 'xxx@xxx.co.jp' -Type scheduled

$tz = [System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time")
$now = Get-Date

$futureMeetings = $Meetings.meetings |
    Where-Object {
        $startUtc = [datetime]$_.start_time
        $startUtc = [datetime]::SpecifyKind($startUtc, [System.DateTimeKind]::Utc)
        $startLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($startUtc, $tz)
        $startLocal -gt $now
    } |
    Sort-Object start_time

Write-Host "=============================="
Write-Host "今後の予定一覧:"
if ($futureMeetings.Count -gt 0) {
    foreach ($m in $futureMeetings) {
        $startUtc = [datetime]$m.start_time
        $startUtc = [datetime]::SpecifyKind($startUtc, [System.DateTimeKind]::Utc)
        $startLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($startUtc, $tz)
        Write-Host ("{0:yyyy-MM-dd HH:mm}({1}分) {2}" -f $startLocal, $m.duration, $m.topic)
    }
} else {
    Write-Host "今後の予定はありません。"
}
Write-Host "=============================="

# --- ミーティング入力 ---
$parsedDate = $null
$duration = 0

# タイトル
$topic = Read-Host "ミーティングのタイトルを入力してください"

# 日付時刻
do {
    $text = Read-Host "ミーティング予定日時を入力してください (例: 2025-01-01 10:00)"
    try {
        $parsedDate = [datetime]::ParseExact($text, 'yyyy-MM-dd HH:mm', $null)
        $result = $true
    } catch {
        Write-Host "日付形式が正しくありません。" -ForegroundColor Red
        $result = $false
    }
} until ($result)

# 時間（分）
do {
    $text = Read-Host "ミーティングの長さ（分）を入力してください"
    [int]::TryParse($text, [ref]$duration)
} until ($duration -ge 30)

# --- 確認 ---
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

# --- ミーティング作成 ---
$meeting = New-ZoomMeeting `
    -UserId 'xxx@xxx.co.jp' `
    -Topic $topic `
    -StartTime $parsedDate.ToString("yyyy-MM-ddTHH:mm:ss") `
    -Duration $duration `
    -Timezone 'Asia/Tokyo'

# --- 結果表示 ---
$ResultText = @"
ミーティングID: $($meeting.id)
タイトル　: $topic
開始日時　: $($meeting.start_time)
時間　　　: $($meeting.duration)分
参加URL 　: $($meeting.join_url)
パスワード: $($meeting.password)
"@

$ResultText | Set-Clipboard
Write-Host "ミーティング作成完了！"
Write-Host $ResultText
Write-Host "ミーティング情報をクリップボードにコピーしました！"
pause
