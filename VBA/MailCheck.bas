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
Set myLocalFolder_admin = myLocalFolder.Folders("admin") '【Custom】ルートフォルダの次の階層のフォルダを指定
Set adminMailItems = myLocalFolder_admin.Items

'メールオブジェクトごとの処理
'件名の取得、カテゴリの判定、返信要否の判定
'一旦配列に格納する
Dim strMsgID As String 'ループ処理用のID
Dim strRpMsgID As String '返信メールのID
Dim objMailItem As Object
Dim intreplystatus As Integer '返信、全員に返信、転送の識別子（102,103,104）
Dim strInterplystatus As String 'Excel出力用の識別子
Dim excelInput() As String 'Excel出力用の多次元配列

Redim excelInput(adminMailitems.count, 5)

'adminフォルダ（Items）内のメール（Item）分だけループ処理
For Each objMailItem In adminMailItems
    intreplystatus = 0 '初期化

    With objMailItem
        strMsgID = .PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        strRpMsgID = adminMailItems.Find("@SQL=""" & PR_IN_REPLY_TO_ID & """ = '" & strMsgID & "'")
        intreplystatus = .PropertyAccessor.GetProperty(PR_LAST_ACTION)



         Select Case intreplystatus
            Case 0
                strInterplystatus = "未返信"
            Case 102
                strInterplystatus = "返信"
            Case 103
                strInterplystatus = "全員に返信"
            Case 104
                strInterplystatus = "転送"
        End Select




    End With









Next objMailItem





'Excelに出力
Dim i As Long 'ループ回数のみを管理する変数
Dim CellCount As Long 'メールの情報を書き出すセルを管理する変数
CellCount = 0

Call 高速化ON

Call 高速化OFF

End Sub