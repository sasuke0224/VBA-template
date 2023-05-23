Attribute VB_Name = "Module7"
Sub mondai6()
    Dim i As Integer
    Dim j As Integer
    j = 1
    For i = 1 To 10
        If Worksheets("Sheet1").Cells(i, 1).Value Mod 2 = 0 Then
        Worksheets("Sheet2").Cells(j, 1).Value = Worksheets("Sheet1").Cells(i, 1).Value
            j = j + 1
        End If
    Next i
End Sub

