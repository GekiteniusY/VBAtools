Sub Excelに出力(ByRef excelOutput() as String)

Range(Range("A1"), Range("A1").Offset(0, excelOutput.count))  = excelOutput


End Sub