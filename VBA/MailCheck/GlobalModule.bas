Dim objOutlook As Outlook.Application
Dim myNamespace As Outlook.Namespace
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"


Sub Initialize()
Set objOutlook = New Outlook.Application
Set myNamespace = objOutlook.GetNamespace("MAPI")   

End Sub