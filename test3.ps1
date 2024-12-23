Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;


public class ToastNotifier
{
    public static void ShowToast(string title, string message)
    {
        string toastXmlString = @"
        <toast>
            <visual>
                <binding template='ToastGeneric'>
                    <text>" + title + @"</text>
                    <text>" + message + @"</text>
                </binding>
            </visual>
        </toast>";
        
        XmlDocument toastXml = new XmlDocument();
        toastXml.LoadXml(toastXmlString);
        ToastNotification toast = new ToastNotification(toastXml);
        ToastNotificationManager.CreateToastNotifier().Show(toast);
    }
}
"@

[ToastNotifier]::ShowToast("Test Title", "This is a test message")