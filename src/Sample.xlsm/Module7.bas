Attribute VB_Name = "Module7"
Sub –â‘è6()
         Dim intLastRowNum As Integer
    Dim i As Integer
    
    With Sheets("Sheet1")
        intLastRowNum = .Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To intLastRowNum
            
            If Worksheets("Sheet1").Cells(i, 1).Value Mod 2 = 0 Then
               Debug.Print .Cells(i, 1).Value
            End If
            
        Next i
        
    End With

    Worksheets("Sheet1").Cells(i, 1).Copy
    Worksheets("Sheet2").Cells(i, 1).PasteSpecial Paste:=xlPasteValues
End Sub

