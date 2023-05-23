Attribute VB_Name = "Module11"
Sub mondai10()
    Dim i As Integer

    For i = 1 To 10
    Worksheets("Sheet2").Cells(i, 1).Value = Worksheets("Sheet1").Cells(i, 1).Value * 10
    Next i

End Sub