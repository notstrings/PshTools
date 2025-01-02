$ErrorActionPreference = "Continue"

switch (
    $Host.UI.PromptForChoice(
        "シャットダウン予約",
        "`tシャットダウン予定時間を選択してください",
        @("Exit(&E)", "1時間(&1)", "3時間(&3)", "5時間(&5)", "8時間(&8)", "シャットダウン取消(&C)"),
        0
    )
) {
    0 { exit 0 }
    1 { shutdown /s /t  3600 /c "ユーザーによるシャットダウン予約" /f /d p:0:0 }
    2 { shutdown /s /t 10800 /c "ユーザーによるシャットダウン予約" /f /d p:0:0 }
    3 { shutdown /s /t 18000 /c "ユーザーによるシャットダウン予約" /f /d p:0:0 }
    3 { shutdown /s /t 18000 /c "ユーザーによるシャットダウン予約" /f /d p:0:0 }
    3 { shutdown /s /t 28800 /c "ユーザーによるシャットダウン予約" /f /d p:0:0 }
    4 { shutdown /a }
}
