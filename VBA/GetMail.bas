Sub GetAdminMail()
Dim objOutlook As Outlook.Application
Dim myNamespace As Outlook.Namespace
Dim myLocalFolder, myLocalFolder_support
Dim Item_support, Item_support_inquiry1, Item_support_inquiry2, Item_support_update
Dim Item_support_rooms, Item_support_obstacle, Item_support_etc

Set objOutlook = New Outlook.Application
Set myNamespace = objOutlook.GetNamespace("MAPI")

Dim i As Long 'ループ回数のみを管理する変数
Dim CellCount As Long 'メールの情報を書き出すセルを管理する変数
CellCount = 0

Dim strFilter As String
strFilter = 受信日時の指定

Set myLocalFolder = myNamespace.Folders("ローカル保存用フォルダ")
Set myLocalFolder_admin = myLocalFolder.Folders("admin")

Set Item_admin = myLocalFolder_admin.Items.Restrict(strFilter)

Call 高速化ON


With ThisWorkbook.Worksheets("admin集計")
    .Cells(1, 2).value = "Title"
    .Cells(1, 3).value = "Category1"
    .Cells(1, 4).value = "Category2"
    .Cells(1, 5).value = "Sender"
    .Cells(1, 6).value = "ReceivedTime"
End With


Call フォルダ毎のループ処理(Item_admin, "admin集計", CellCount)


Call 高速化OFF


End Sub