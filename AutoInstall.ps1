
Write-Host "自動構成中"
winget configure "$($PSScriptRoot)\DSC\General.dsc"
winget configure "$($PSScriptRoot)\DSC\Develop.dsc"

# NOTE
# PhotoScape
# XTRMRUntime
# https://boonx4m312s.hatenablog.com/entry/2023/05/10/180000

# 管理権限必要
# "7zip.7zip"
# "ImageMagick.ImageMagick.Q16"
# "UB-Mannheim.TesseractOCR"
# "WiresharkFoundation.Wireshark"
# "JohnMacFarlane.Pandoc"
