Sub AdminFlagging()

'------------READ ME------------
'コメント内で【Custom】と付与されている箇所は
'自身の環境に応じて変える必要があります。
'
'
'-------------------------------

'変数の宣言
Dim objOutlook As Outlook.Application
Dim myNamespace As Outlook.Namespace
Dim myLocalFolder, myLocalFolder_admin As Folder
Dim adminMailItems As Object
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
Dim strMsgID As String
Dim objMailItem As Object
Dim intreplystatus As Integer

For Each objMailItem In adminMailItems

    With objMailItem



    End With

Next objMailItem





'Excelに出力
Dim i As Long 'ループ回数のみを管理する変数
Dim CellCount As Long 'メールの情報を書き出すセルを管理する変数
CellCount = 0

Call 高速化ON

Call 高速化OFF

End Sub