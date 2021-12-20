Sub AdminFlagging()
'------------READ ME------------
'コメント内で【Custom】と付与されている箇所は
'自身の環境に応じて変える必要があります。
'
'Outlookオブジェクト操作用の共通設定----------------------------------------------------------------------------
Dim objOutlook As Outlook.Application
Dim myNamespace As Outlook.Namespace
Set objOutlook = New Outlook.Application
Set myNamespace = objOutlook.GetNamespace("MAPI")
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"


'メールボックス・メールアイテム操作用の個別設定---------------------------------------------------------------------
Dim myLocalFolder, myLocalFolder_admin As Folder
Dim adminMailItems As Object
Set myLocalFolder = myNamespace.Folders("ローカル保存用フォルダ") '【Custom】ローカルのルートフォルダを指定
Set myLocalFolder_admin = myLocalFolder.Folders("admin")        '【Custom】ルートフォルダの次の階層のフォルダを指定

Dim strFilter As String
strFilter = 受信日時の指定
Set adminMailItems = myLocalFolder_admin.Items.Restrict(strFilter)


'----------------------------------------------------------------------------------------------------------------
'１．メールオブジェクトごとにフラグ付け
'２．フラグ付けの結果、件名、受信日時は配列に格納

'フラグ付けの結果などを格納するための配列：Excel出力に利用
Dim excelOutput() As String                     
ReDim excelOutput(adminMailItems.Count + 50, 2) 'カテゴリ分け実装時には2→4に変更

'返信メールの有無をチェックするための変数
Dim strMsgID As String, strRpMsgID As MailItem, strInterplystatus As String
'Dim tag as String 'カテゴリ分け実装用

'For Eachループで使用
Dim objMailItem As Object, i As Integer: i = 0

For Each objMailItem In adminMailItems  'adminフォルダ（Items）内のメール（Item）分だけループ処理
    With objMailItem
    
        'tag = メールの分類(objMailItem)    'カテゴリ分け実装用

        Select Case True    'RE,Re,FWがついている
            Case .Subject Like "*RE*"
                strInterplystatus = "OK"
            Case .Subject Like "*Re*"
                strInterplystatus = "OK"
            Case .Subject Like "*FW*"
                strInterplystatus = "OK"
        End Select

        If strInterplystatus = "OK"  Then 
            End If
        Else 'REがついていない
            strMsgID = .PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
            Set strRpMsgID = adminMailItems.Find("@SQL=""" & PR_IN_REPLY_TO_ID & """ = '" & strMsgID & "'")
  
            If strRpMsgID Is Nothing Then
                strInterplystatus = "未返信"
            Else
                strInterplystatus = "OK"
            End If
        End If        

        excelOutput(i, 0) = .ReceivedTime
        excelOutput(i, 1) = .Subject
        excelOutput(i, 2) = strInterplystatus
        'excelOutput(i, 3) = .Body  カテゴリ分け実装用
        'excelOutput(i, 4) = tag  カテゴリ分け実装用

        '初期化
        'tag = ""
        i = i + 1

    End With

Next objMailItem


Call 高速化ON
Call Excelに出力(excelOutput)
Call 高速化OFF

End Sub