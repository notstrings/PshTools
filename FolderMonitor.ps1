$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Invoke-Expression -Command @"
    class FolderMonitorConf {
        [MonitorTarget[]] `$MonitorTargets
        [System.ComponentModel.Description("変更の適用に再起動が必要")]
        [int] `$Interval
    }
    class MonitorTarget {
        [System.ComponentModel.Description("監視名称")]
        [string] `$MonName
        [System.ComponentModel.Description("監視位置")]
        [string] `$MonPath
    }
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $child1 = New-Object MonitorTarget -Property @{MonName = "Mon01"; MonPath = ""}
        $child2 = New-Object MonitorTarget -Property @{MonName = "Mon02"; MonPath = ""}
        $child3 = New-Object MonitorTarget -Property @{MonName = "Mon03"; MonPath = ""}
        $Conf = New-Object FolderMonitorConf -Property @{
            MonitorTargets = @($child1, $child2, $child3)
            Interval = (5*60*1000)
        }
        SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [FolderMonitorConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([FolderMonitorConf]) $json
    return $Conf
}
# 設定編集
function local:EditConfFile([string] $Title, [string] $Path) {
    $Conf = LoadConfFile $Path
    $ret = ShowSettingDialog $Title $Conf
    if ($ret -eq "OK") {
        SaveConfFile $Path $Conf
    }
}

## 本体 #######################################################################

function local:FolderMonitor() {
    # 設定取得
    $Conf = LoadConfFile $ConfPath
    # 更新検出
    $Result = @()
    if ($null -ne $Conf.MonitorTargets) {
        $Conf.MonitorTargets | ForEach-Object {
            $MonitorName = $_.MonName
            $MonitorPath = $_.MonPath
            if ( ("" -ne $MonitorPath) -and (Test-Path -LiteralPath $MonitorPath)) {
                $Result += CheckFolderUpdate $MonitorName $MonitorPath
            }
        }
    }
    # 結果表示
    if ($Result.Count -gt 0){
        # ここでブロッキングされても変更は次の処理で見つかる
        ShowUpdateFileList -Title "フォルダ監視結果" -Message "以下のファイルが追加/削除/変更されました。" -Results $Result
    }
}
function local:CheckFolderUpdate([string] $MonitorName, [string] $MonitorPath) {
    $Ret = @()

    # ワーキングフォルダを確保
    $PrevPath = "$($PSScriptRoot)\Monitor\$($MonitorName)Prev.csv"
    $CrntPath = "$($PSScriptRoot)\Monitor\$($MonitorName)Crnt.csv"
    $null = New-Item ([System.IO.Path]::GetDirectoryName($PrevPath)) -ItemType Directory -ErrorAction SilentlyContinue
    $null = New-Item ([System.IO.Path]::GetDirectoryName($CrntPath)) -ItemType Directory -ErrorAction SilentlyContinue

    # 現在の監視フォルダ状況を取得
    Get-ChildItem $MonitorPath -File -Recurse |
    Select-Object FullName, LastWriteTime |
    Export-Csv -Path $CrntPath -NoTypeInformation -Encoding "OEM" # PowerShell5互換のためエンコーディングを文字指定

    # 昨今の監視フォルダ状況差分を確認
    if (Test-Path $PrevPath){
        # ファイル名と変更日時の差分を抽出
        $PrevAllCSV = @(Import-Csv -LiteralPath $PrevPath -Encoding "OEM") # PowerShell5互換のためエンコーディングを文字指定
        $CrntAllCSV = @(Import-Csv -LiteralPath $CrntPath -Encoding "OEM") # PowerShell5互換のためエンコーディングを文字指定
        $PrevDifCSV = @()
        $CrntDifCSV = @()
        Compare-Object -ReferenceObject $PrevAllCSV -DifferenceObject $CrntAllCSV -Property FullName, LastWriteTime |
        ForEach-Object {
            if($_.SideIndicator -eq "<=") {
                $PrevDifCSV += [PSCustomObject]@{FullName = $_.FullName; LastWriteTime = $_.LastWriteTime}
            } elseif ($_.SideIndicator -eq "=>") {
                $CrntDifCSV += [PSCustomObject]@{FullName = $_.FullName; LastWriteTime = $_.LastWriteTime}
            }
        } | Out-Null
        # ファイル名差分だけで絞り込んで追加/削除/変更を識別
        $ModFile = @()
        $DelFile = @()
        $AddFile = @()
        Compare-Object -ReferenceObject $PrevDifCSV -DifferenceObject $CrntDifCSV -IncludeEqual -Property FullName |
        ForEach-Object {
            if($null -ne $_.FullName -and "FullName" -ne $_.FullName){
                if($_.SideIndicator -eq "<=") {
                    $DelFile += $_.FullName
                } elseif ($_.SideIndicator -eq "=>") {
                    $AddFile += $_.FullName
                } elseif ($_.SideIndicator -eq "==") {
                    $ModFile += $_.FullName
                }
            }
        } | Out-Null
        # 結果出力
        $AddFile | ForEach-Object { $Ret += [PSCustomObject]@{MonitorName = $MonitorName; MonitorPath = $MonitorPath; Type = "ADD"; Path=$_ } } | Out-Null
        $DelFile | ForEach-Object { $Ret += [PSCustomObject]@{MonitorName = $MonitorName; MonitorPath = $MonitorPath; Type = "DEL"; Path=$_ } } | Out-Null
        $ModFile | ForEach-Object { $Ret += [PSCustomObject]@{MonitorName = $MonitorName; MonitorPath = $MonitorPath; Type = "MOD"; Path=$_ } } | Out-Null
    }

    # 現在のフォルダ状況を過去のフォルダ状況とする
    $null = Copy-Item -Destination $PrevPath -LiteralPath $CrntPath -Force

    return $Ret
}
function local:ShowUpdateFileList(
    [Parameter(Mandatory = $true)] [string]$Title,
    [Parameter(Mandatory = $true)] [string]$Message,
    [Parameter(Mandatory = $false)] [object[]]$Results
)
{
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    # 画面生成
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Dialog" Width="480" Height="320" WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <DockPanel>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="5">
            <Button Name="btnOK" Content="OK" Width="128" Height="50"/>
        </StackPanel>
        <Label Name="lblMessage" DockPanel.Dock="Top"/>
        <DockPanel>
            <DataGrid Name="grdFiles" ItemsSource="{Binding Files}"
                AutoGenerateColumns="False" AllowDrop="False" SelectionMode="Single"
                CanUserAddRows="False" CanUserDeleteRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Monitor" Binding="{Binding MonitorName}" Width="*" IsReadOnly="True"/>
                    <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="*" IsReadOnly="True"/>
                    <DataGridTextColumn Header="Name" Binding="{Binding Path}" Width="3*" IsReadOnly="True"/>
                </DataGrid.Columns>
            </DataGrid>
        </DockPanel>
    </DockPanel>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $lblMessage = $window.FindName("lblMessage")
    $grdFiles   = $window.FindName("grdFiles")
    $btnOK      = $window.FindName("btnOK")

    # Window
    $window.Title = $Title
    $lblMessage.Content = $Message
    $bndFiles = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    foreach ($Result in $Results) {
        $bndFiles.Add($Result)
    }
    $DataContext = [PSCustomObject]@{
        Files = $bndFiles
    }
    $window.DataContext = $DataContext

    # DataGrid
    $grdFiles.Add_MouseDoubleClick({
        $Result = $grdFiles.SelectedItem
        if ($null -eq $Result) {
            return
        }
        $Path = $Result.Path
        while( (Test-Path -LiteralPath $Path) -eq $false ) {
            $Path = [System.IO.Path]::GetDirectoryName($Path)
            if ([string]::IsNullOrEmpty($Path)) {
                return
            }
        }
        Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
    })

    # OK
    $btnOK.Add_Click({
        $window.DialogResult = $true
    })

    # 表示
    $window.ShowDialog() | Out-Null
}

###############################################################################

try {
    $null = Write-Host "---$Title---"
    InitConfFile $ConfPath
    $Conf = LoadConfFile $ConfPath
    RunInTaskTray $Title 0x0000ff { EditConfFile $Title $ConfPath } { FolderMonitor } $Conf.Interval
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
