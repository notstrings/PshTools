# オプション指定が面倒臭い
function local:ExecWinget([string] $PackageID) {
    Write-Host "winget install --id $PackageID -e --accept-package-agreements --source winget"
    winget install --id $PackageID -e --accept-package-agreements --source winget
    Write-Host ""
}

# 管理権限必要
# ExecWinget "7zip.7zip"
# ExecWinget "ImageMagick.ImageMagick.Q16"
# ExecWinget "UB-Mannheim.TesseractOCR"
# ExecWinget "WiresharkFoundation.Wireshark"
# ExecWinget "JohnMacFarlane.Pandoc"

# 管理権限不要
ExecWinget "Microsoft.VisualStudioCode"
ExecWinget "Microsoft.PowerToys"
ExecWinget "Git.Git"
ExecWinget "Atlassian.Sourcetree"
ExecWinget "DevToys-app.DevToys"
ExecWinget "FastCopy.IPMsg"

# NOTE
# PhotoScape
# XTRMRUntime
# https://boonx4m312s.hatenablog.com/entry/2023/05/10/180000
