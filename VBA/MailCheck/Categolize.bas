Function メールの分類(ByVal mailItem as Object)
    'Status：作成途中
    '出力先シートの指定
Dim mailBody as String, mailTitle as String
Dim searchWord as String

mailTitle = mailItem.Subject
mailBody = mailItem.Body

'集計表をフィルタ===========================================
if InStr(mailTitle, "RE") = 0 Then
    searchWord = Mid(mailTitle,Len("(admin XXXXXX)  ")+1)  'REが含まれていない
Else
    searchWord = Mid(mailTitle,Len("(admin 295751)  RE: ")+1)  'REが含まれている
End if

'２．Flag列にカテゴリが立っているかを確認
'集計表をフィルタ===========================================

'３－０：すべてにフラグが立っている場合
'何もしない

'３－１：１つ立っている場合
'そのフラグをコピーして、ほかの列に貼り付け


'３－２：２つ以上たっている場合
'一番上のフラグをコピーして、ほかの列に貼り付け


'３－３：何もたっていない場合
'　３－３－０：件名に"RE: "、"FW: "のどちらも含まれていない
'	件名、本文からカテゴリ分け
'　３－３－１：件名に"RE: "または"FW: "が含まれている
'	何もしない

'集計表のフィルタ解除===========================================
If searchWord = vbNullString Then                                   '入力ない場合は処理終了'
    Exit Sub
Else
    ActiveSheet.Range("A1").AutoFilter 4, "*" & searchWord & "*"   '検索文字列を含むセルで検索'
End If


End Function