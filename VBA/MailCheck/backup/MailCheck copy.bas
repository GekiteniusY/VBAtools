Sub AdminFlagging()

'------------READ ME------------
'コメント内で【Custom】と付与されている箇所は
'自身の環境に応じて変える必要があります。
'
'
'
'-------------検討中-------------
'For Eachでループ処理をさせる都合上、処理が終わったメールはフォルダを移す
'運用にするべきか？それとも、配列に格納する際に、日付指定をするか。
'1日のメール量を考慮して、ハイブリッド型がよさそう
'-------------------------------

'変数の宣言
Dim objOutlook As Outlook.Application
Dim myNamespace As Outlook.Namespace
Dim myLocalFolder, myLocalFolder_admin As Folder
Dim adminMailItems As Object
Const PR_LAST_ACTION = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001E"

'メールオブジェクトの取得、フォルダの設定
Set objOutlook = New Outlook.Application
Set myNamespace = objOutlook.GetNamespace("MAPI")
Set myLocalFolder = myNamespace.Folders("ローカル保存用フォルダ") '【Custom】ローカルのルートフォルダを指定
Set myLocalFolder_admin = myLocalFolder.Folders("admin")        '【Custom】ルートフォルダの次の階層のフォルダを指定

Dim strFilter As String
strFilter = 受信日時の指定
Set adminMailItems = myLocalFolder_admin.Items.Restrict(strFilter)

'メールオブジェクトごとの処理
'件名の取得、カテゴリの判定、返信要否の判定
'一旦配列に格納する
Dim strMsgID As String, strRpMsgID As MailItem  '返信メールの有無をチェックするための変数
Dim objMailItem As Object
Dim intreplystatus As Integer   '返信、全員に返信、転送の識別子（102,103,104）
Dim strInterplystatus As String '返信有無の識別子
Dim excelOutput() As String      'Excel出力用の多次元配列
Dim i As Integer: i = 0         '配列に格納するための変数
ReDim excelOutput(adminMailItems.Count + 50, 4)


'Dim tag as String 'カテゴリ分け実装用

For Each objMailItem In adminMailItems  'adminフォルダ（Items）内のメール（Item）分だけループ処理
    intreplystatus = 0 '初期化

    With objMailItem

        '返信フラグの情報を取得：strInterplystatus
        'intreplystatus = .PropertyAccessor.GetProperty(PR_LAST_ACTION)
        'Select Case intreplystatus
        '   Case 0
        '       strInterplystatus = "未返信"
        '   Case 102
        '       strInterplystatus = "返信"
        '   Case 103
        '       strInterplystatus = "全員に返信"
        '   Case 104
        '       strInterplystatus = "転送"
        'End Select

        'メール部の分類情報を取得
        'tag = メールの分類(objMailItem)

        
        
        If .Subject  Like "*RE*"  Then 'REがついている
            strInterplystatus = "OK"
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