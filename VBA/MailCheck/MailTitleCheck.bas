Function ReplyCheck(ByVal mailTitle as String) As String
    'Status：完成
    '修正予定なし

        Select Case True    'RE,Re,FWがついている
            Case mailTitle Like "*RE*"
                ReplyCheck = "OK"
            Case mailTitle Like "*Re*"
                ReplyCheck = "OK"
            Case mailTitle Like "*FW*"
                ReplyCheck = "OK"
            Case Else
                ReplyCheck = "未返信"
        End Select
End Function

Function TitleCategoryCheck(ByVal mailTitle as String) As String
    'Status：未完成
    '作成途中


    Dim tag As String

    If tag = "" Then

    Else
        tag = "不明"
    End If

    TitleCategoryCheck = tag

 '検索にヒットしない場合は""を返すこと
End Function

Function TitleCheckAtSummaryFile(ByVal mailTitle as String) As String
    'Status：未完成
    '作成途中

    Dim tag As String

 '検索にヒットしない場合は""を返すこと
End Function


Function BodyCategoryCheck(ByVal mailBody as String) As String
    'Status：未完成
    '作成途中

    Dim tag As String

    BodyCategoryCheck = ""
    
    'このファンクションで判定できない場合は、"判定不可"のフラグを付ける
End Function