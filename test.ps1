$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$result = ShowFileListDialog -Title "ファイルを選択してください" -Message "ここにファイルをドラッグ＆ドロップ" -FileFilter ".*" -Files $args
if ($result.Result) {
    foreach ($file in $result.Files) {
        Write-Host $file
    }
}

# OpenCV sample PowerShell 5.1
# nuget install OpenCvSharp3-AnyCPU
# $opencv = Join-Path $PSScriptRoot "tool\OpenCvSharp3-AnyCPU.4.0.0.20181129\lib\net461\OpenCvSharp.dll"
# $native = Join-Path $PSScriptRoot "tool\OpenCvSharp3-AnyCPU.4.0.0.20181129\runtimes\win10-x64\native"
# $env:PATH = "$native;$env:PATH"
# [System.Reflection.Assembly]::LoadFrom($opencv)
# $img = New-Object OpenCvSharp.Mat(128,128,[OpenCvSharp.MatType]::CV_8UC1)
# [OpenCvSharp.Cv2]::ImWrite("$PSScriptRoot\test.png",$img)
# $img.Dispose()
