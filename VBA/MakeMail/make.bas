Sub メール作成()




'＝＝＝＝＝＝＝＝＝＝＝＝＝＝用意する変数＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝




'メールタイトルのString
Dim title As String
Set title = "メールタイトル" '======================================================================================タイトル

'メール本文のString
Dim body As String

'メールの雛形のString
Dim Template As String

'CCの宛先のString
'Excelから抽出した情報を一時格納するMap
Dim map As Object
Set map = CreateObject("Scripting.Dictionary")
'メールアイテムの保存先
Dim myNamespace As Outlook.Namespace
Set myNamespace = objOutlook.GetNamespace("MAPI") 
Dim draftFolder As Folder
Set draftFolder = myNamespace.Folders("yogi.genkichi.za@ncontr.com").Folders("下書き") 
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝


'Excelから情報を抽出してMapに格納する処理
Call map.Add("To", "送信先")
Call map.Add("CC", "Cc")
Call map.Add("Body", "本文")



'Listの情報をもとにメールを作成する処理
Dim i As Integer
i = 4
Do While i = 0
    Dim Outlook_obj As Outlook.Application
    Dim MailItem_obj As Oulook.MailItem
    Set Outlook_obj = CreateObject("Outlook.Application")
    Set MailItem_obj = Outlook_obj.CreateItem(olMailItem)

    With MailItem_obj
        .To = map.Item("To")
        .CC = map.Item("CC")
        .Subject = title
        .Body = map.Item("Body")
    End With

    

    i = i - 1
Loop





'作成したメールオブジェクトをフォルダに保存する処理










Dim adminMailItems As Object

'----------------------------------------------------------------------------------------------------------------

Dim excelOutput() As String
ReDim excelOutput(adminMailItems.Count + 50, 3)


Dim mailTitle As String
Dim strInterplystatus As String
Dim strMsgID As String, strRpMsgID As MailItem
Dim tag As String
tag = ""

'For Eachループで使用
Dim objMailItem As Object, i As Integer: i = 0


For Each objMailItem In adminMailItems
    With objMailItem

        mailTitle = .Subject
        strMsgID = .PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        Set strRpMsgID = adminMailItems.Find("@SQL=""" & PR_IN_REPLY_TO_ID & """ = '" & strMsgID & "'")

        '返信済みかどうかのチェック===========================================
        If strRpMsgID Is Nothing Then
            strInterplystatus = ReplyCheck(mailTitle)   '親アイテムが見つからない場合：タイトルで判定
        Else
            strInterplystatus = "OK"                    '親アイテムがある場合：OKフラグ
        End If


        '配列への格納===========================================
        excelOutput(i, 0) = .ReceivedTime
        excelOutput(i, 1) = mailTitle
        excelOutput(i, 2) = strInterplystatus
        excelOutput(i, 3) = tag

        '初期化とインクリメント===========================================
        mailTitle = ""
        strInterplystatus = ""
        tag = ""
        i = i + 1

    End With

Next objMailItem


Call 高速化ON
Call Excelに出力(excelOutput)
Call 高速化OFF







End Sub
