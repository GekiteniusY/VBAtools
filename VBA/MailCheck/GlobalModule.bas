Public objOutlook As Outlook.Application
Public myNamespace As Outlook.Namespace
Public Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Public Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"


Sub Initialize()
    Set objOutlook = New Outlook.Application
    Set myNamespace = objOutlook.GetNamespace("MAPI")   S
End Sub