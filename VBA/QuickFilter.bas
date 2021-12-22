Sub FilterON()
'フィルタ機能を利用するだけのマクロです'
'フィルタにかける列に文字列が多く格納されている場合、GUI上では表示に時間がかかるため作成しました'

Dim searchWord As String

 On Error GoTo ErrorHandler
    searchWord = InputBox("検索文字列を入力してください。")            '検索文字列の入力'
    
    If searchWord = vbNullString Then                                   '入力ない場合は処理終了'
        Exit Sub
    Else
        Call cellState(ActiveCell.Address(False, False))                'アクティブセルの値を保存'
        ActiveSheet.Range("A1").AutoFilter 1, "*" & searchWord & "*"   '検索文字列を含むセルで検索'
    End If
Exit Sub

ErrorHandler:
        MsgBox ("エラーが発生しました")

End Sub



Sub FilterOFF()
    Dim cell As String
    ActiveSheet.ShowAllData  'フィルタ解除'
    cell = cellState("get")

    If cell = vbNullString Then
    Else
        Range(cell).Activate    'アクティブセルを復元'
    End If
    
    Call cellState("clear")             'アクティブセルの値を初期化'
End Sub



Private Function cellState(state As String) As String
'FilterONプロシージャで呼び出された際はアクティブセルの値を保存'
'FilterOFFプロシージャで呼び出された際は保存したアクティブセルの値を返す'
'VBにstaticなグローバル変数が存在しないため作成したものです'
   
    Static staticCellState As String
    
    If (StrPtr(staticCellState) = 0) Then   '初回起動or初期化済であればアクティブセルの値を保存'
        staticCellState = state
        Exit Function
    ElseIf state = "get" Then               '引数がgetのとき＝FilterOFFで1回目の呼び出しのとき、保存したセルの状態を返す'
        cellState = staticCellState
    ElseIf state = "clear" Then             '引数がclearのとき＝FilterOFFで2回目の呼び出しのとき、変数を初期化'
        staticCellState = vbNullString
        Exit Function
    Else                                    'FilterONで呼び出した場合はここで処理を抜ける;戻り値を返す必要がないため'
        Exit Function
    End If
    
End Function