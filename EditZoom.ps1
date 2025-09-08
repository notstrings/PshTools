# =========================================
# Zoomミーティングスケジュール作成スクリプト

# --- モジュール確認 ---
if (-not (Get-Module -ListAvailable -Name PSZoom)) {
    Install-Module -Name PSZoom -Scope CurrentUser -Force -AllowClobber
}
Import-Module PSZoom

# --- 認証情報 ---
$sCnfPath = ".\Config\EditZoom.json"
if (-not (Test-Path $sCnfPath)) {
    Write-Host "設定ファイルが見つかりません: $sCnfPath" -ForegroundColor Red
    exit 1
}
$Cnf = Get-Content $sCnfPath -Raw | ConvertFrom-Json

# 予定表示
function local:ShowMeetingSchedule() {
    Connect-PSZoom -AccountID $Cnf.AccountID -ClientID $Cnf.ClientID -ClientSecret $Cnf.ClientSecret
    $Meetings = `
        (Get-ZoomMeetingsFromuser -UserId $Cnf.UserID -Type scheduled).meetings |
        Where-Object {
            $STime = ([System.TimeZoneInfo]::ConvertTimeFromUtc($_.start_time, [System.TimeZoneInfo]::Local))
            $CTime = (Get-Date)
            $CTime -le $STime 
        } |
        Sort-Object start_time
    Write-Host "=============================="
    Write-Host "今後の予定一覧:"
    foreach ($m in $Meetings) {
        $STime = ([System.TimeZoneInfo]::ConvertTimeFromUtc($m.start_time, [System.TimeZoneInfo]::Local))
        Write-Host (" {0} | {1:yyyy-MM-dd HH:mm}({2}分) {3}" -f $m.id, $STime, $m.duration, $m.topic)
    }
    Write-Host "=============================="
}

# 予定追加
function local:AddMeetingSchedule() {
    # --- 入力 ---
    $Topic  = ""
    $StartTime = $null
    $Duration = 0

    # タイトル
    $Topic = Read-Host "ミーティングのタイトルを入力してください"
    if($Topic -eq "") {
        return
    }

    # 日付時刻
    do {
        $text = Read-Host "ミーティング予定日時を入力してください (例:2025-01-01 17:00)"
        if($text -eq "") {
            return
        }
        try {
            $StartTime = [datetime]::ParseExact($text, 'yyyy-MM-dd HH:mm', $null)
            $result = $true
        } catch {
            Write-Host "日付形式が正しくありません。" -ForegroundColor Red
            $result = $false
        }
    } until ($result)

    # 時間(分)
    do {
        $text = Read-Host "ミーティングの長さ(分)を入力してください"
        if($text -eq "") {
            return
        }
        [int]::TryParse($text, [ref]$Duration)
    } until ($Duration -ge 30)

    # --- 確認 ---
    Write-Host "=============================="
    Write-Host "以下の内容でミーティングを作成します"
    Write-Host "タイトル : $Topic"
    Write-Host "開始日時 : $StartTime"
    Write-Host "時間     : $Duration 分"
    Write-Host "=============================="
    $text = Read-Host "実行しますか？ (Y/N)"
    if ($text -ne 'Y' -and $text -ne 'y') {
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        return
    }

    # --- 作成 ---
    Connect-PSZoom -AccountID $Cnf.AccountID -ClientID $Cnf.ClientID -ClientSecret $Cnf.ClientSecret
    $Meeting = New-ZoomMeeting `
        -UserId $Cnf.UserID `
        -Topic $Topic `
        -StartTime $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") `
        -Duration $Duration

    # --- 結果 ---
    $text = ""
    $text = $text + "ミーティングID: $($Meeting.id)`n"
    $text = $text + "タイトル      : $($Meeting.topic)`n"
    $text = $text + "開始日時      : $(([System.TimeZoneInfo]::ConvertTimeFromUtc($Meeting.start_time, [System.TimeZoneInfo]::Local)))`n"
    $text = $text + "時間          : $($Meeting.duration)分`n"
    $text = $text + "参加URL       : $($Meeting.join_url)`n"
    $text = $text + "パスワード    : $($Meeting.password)`n"
    Write-Host "ミーティング作成完了！"
    Write-Host $text

    # --- 追加 ---
    $text | Set-Clipboard
    Write-Host "ミーティング情報をクリップボードにコピーしました"
}

# 予定削除
function local:RemoveMeetingSchedule() {
    # --- 入力 ---
    $MeetingID  = ""

    # タイトル
    $MeetingID = Read-Host "ミーティングのIDを入力してください"
    if($MeetingID -eq "") {
        return
    }

    # --- 確認 ---
    Write-Host "以下のミーティングを削除します"
    Write-Host "ID : $MeetingID"
    $text = Read-Host "実行しますか？ (Y/N)"
    if ($text -ne 'Y' -and $text -ne 'y') {
        Write-Host "キャンセルしました。" -ForegroundColor Yellow
        return
    }

    # --- 削除 ---
    Connect-PSZoom -AccountID $Cnf.AccountID -ClientID $Cnf.ClientID -ClientSecret $Cnf.ClientSecret
    Remove-ZoomMeeting -meeting_id $MeetingID 
}

# 本体
while ($true) {
    switch (
        $Host.UI.PromptForChoice(
            "Zoomミーティングスケジュール編集",
            "コマンドを選択してください",
            @("Exit(&E)", "予定表示(&V)", "予定作成(&A)", "予定削除(&D)"),
            0
        )
    ) {
        0 { exit 0 }
        1 { ShowMeetingSchedule }
        2 { AddMeetingSchedule }
        3 { RemoveMeetingSchedule }
    }
}
