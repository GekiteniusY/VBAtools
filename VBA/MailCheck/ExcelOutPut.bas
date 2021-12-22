Sub Excelに出力(excelOutput() As String)
    'Status：一部完成
    'カテゴリわけの情報の出力に対応させる
Dim i As Integer
Dim ws As Worksheet

Set ws = Worksheets("マクロ出力先")
i = UBound(excelOutput) - LBound(excelOutput) - 1

ws.Range(Range("A2"), Range("A2").Offset(i, 3)) = excelOutput

End Sub