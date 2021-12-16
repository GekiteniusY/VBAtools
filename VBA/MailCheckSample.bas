Sub anotherLogic()
' ここをトリプルクリックでマクロ全体を選択できます。
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"
'
Sub FindRepliedItem()
    Dim objItem As MailItem
    Dim strMsgID As String
    Dim fldSent As Folder
    Dim colItems As Items
    Dim oneItem As MailItem


    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Set objItem = ActiveInspector.CurrentItem
    Else
        Set objItem = ActiveExplorer.Selection(1)
    End If
    '
    strMsgID = objItem.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    Set fldSent = Session.GetDefaultFolder(olFolderSentMail) '送信済みフォルダを指定
    Set colItems = fldSent.Items
    Set oneItem = colItems.Find("@SQL=""" & PR_IN_REPLY_TO_ID & """ = '" & strMsgID & "'")


    If oneItem Is Nothing Then
        MsgBox "返信アイテムが見つかりませんでした。"
    Else
        oneItem.Display
        Set oneItem = colItems.FindNext
        While Not oneItem Is Nothing
            oneItem.Display
            Set oneItem = colItems.FindNext
        Wend
    End If



End Sub
