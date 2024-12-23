

function ShowToast {
    param (
        [Parameter(Mandatory = $true)] [string] $Title, 
        [Parameter(Mandatory = $true)] [string] $Message, 
        [Parameter(Mandatory = $true)] [string] $Detail
    )
    begin {}
    process {
        $blk = {
            function ShowToastInner($Title, $Message, $Detail) {
                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
                [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
                [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
                $app_id = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
                $content = ""
                $content += "<?xml version='1.0' encoding='utf-8'?>"
                $content += "<toast>"
                $content += "<visual>"
                $content += "    <binding template='ToastGeneric'>"
                $content += "        <text>$Title</text>"
                $content += "        <text>$Message</text>"
                $content += "        <text>$Detail</text>"
                $content += "    </binding>"
                $content += "</visual>"
                $content += "</toast>"
                $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
                $xml.LoadXml($content)
                $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
                [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app_id).Show($toast)
            }
            ShowToastInner __ARG1 __ARG2 __ARG3
        }
        $enctxt = $blk.ToString()
        $enctxt = $enctxt.Replace("__ARG1", $Title)
        $enctxt = $enctxt.Replace("__ARG2", $Message)
        $enctxt = $enctxt.Replace("__ARG3", $Detail)
        $enccmd = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($enctxt))
        Start-Process "cmd.exe" -ArgumentList "/c start /min powershell -NoProfile -EncodedCommand $enccmd"
    }
end {}
}

ShowToast "aaa" "bbb" "ccc"
