$ErrorActionPreference = "Continue"

function local:ExecWinget([string] $PackageID) {
    Write-Host "winget install --id $PackageID -e --accept-package-agreements --source winget"
    winget install --id $PackageID -e --accept-package-agreements --source winget
    Write-Host ""
}

switch (
    $Host.UI.PromptForChoice(
        "シャットダウン予約",
        "`tシャットダウン予定時間を選択してください",
        @("Exit(&E)", "1時間(&1)", "3時間(&3)", "5時間(&5)", "シャットダウン取消(&C)"),
        0
    )
) {
    0 { exit 0               }
    1 { shutdown -s -t  3600 }
    2 { shutdown -s -t 10800 }
    3 { shutdown -s -t 18000 }
    4 { shutdown -a          }
}

switch (
    $Host.UI.PromptForChoice(
        "画面ロック抑止",
        "`t画面ロック中はシャットダウンされません`n`t画面ロック回避ツールをインストールしますか？",
        @("NO(&N)","YES(&Y)"),
        0
    )
) {
    0 { exit 0 }
    1 {
        Write-Host "自動構成中"
        winget install --id "Microsoft.PowerToys" -e --accept-package-agreements --source winget
        winget configure .\PowerToys.dsc
    }
}
