Sub AdminFlagging()
'------------READ ME------------
'コメント内で【Custom】と付与されている箇所は
'自身の環境に応じて変える必要があります。
'
'Outlookオブジェクト操作用の共通設定----------------------------------------------------------------------------
Call Initialize


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
ReDim excelOutput(adminMailItems.Count + 50, 3) 

'チェック用の変数
Dim mailTitle As String
Dim strInterplystatus As String
Dim strMsgID As String, strRpMsgID As MailItem
Dim tag as String
tag = ""

'For Eachループで使用
Dim objMailItem As Object, i As Integer: i = 0


For Each objMailItem In adminMailItems                  'adminフォルダ（Items）内のメール（Item）分だけループ処理
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

        '問い合わせのカテゴリ判定===========================================
        If strRpMsgID Is Nothing Then
            tag = TitleCategoryCheck(mailTitle)         '親アイテムが見つからない場合：タイトルでカテゴリ判定
        Else
            tag = TitleCheckAtSummaryFile(mailTitle)    '親アイテムがある場合：集計元のExcelデータでカテゴリ判定
        End If
        
        If tag = "" The                                 'タイトルで判定がつかない且つ、集計元のExcelにも情報がない場合
            tag = BodyCategoryCheck(.body)              'メールの本文でカテゴリ判定
        End If

        '配列への格納===========================================
        excelOutput(i, 0) = .ReceivedTime
        excelOutput(i, 1) = mailTitle
        excelOutput(i, 2) = strInterplystatus
        excelOutput(i, 4) = tag

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