Sub 返信転送日チェッカー2()

'参照設定を行い、OutlookApplicationインスタンスを作成
  Dim appOL As Outlook.Application
  Set appOL = New Outlook.Application

'受信トレイのメールをすべて抽出
  Dim objItems As Object
  Set objItems = appOL.GetNamespace("MAPI").GetDefaultFolder(6).Items
  
  Dim i As Integer: i = 2

  Dim objMailItem As Object 'ForEach用の変数。メールアイテムコレクションからメールアイテムを代入
  Dim intreplystatus As Integer '102=返信、103=全員に返信、104=転送、それ以外=0
  
  For Each objMailItem In objItems
  
    With objMailItem
    
      'プロパティタグを使う。0x10810003が返信・全員に返信・転送されたかのプロパティを示す
      intreplystatus = .PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
      
      If intreplystatus <> 0 Then
            
        Cells(i, 1).Value = .ReceivedTime
        Cells(i, 2).Value = .Subject
        Cells(i, 3).Value = .SenderName
        Cells(i, 4).Value = .LastModificationTime
        
        If intreplystatus = 102 Then
          Cells(i, 5).Value = "返信"
        ElseIf intreplystatus = 103 Then
          Cells(i, 5).Value = "全員に返信"
        ElseIf intreplystatus = 104 Then
          Cells(i, 5).Value = "転送"
        End If
          
        intreplystatus = 0
        
        i = i + 1
      End If
      
    End With
    
  Next objMailItem

End Sub