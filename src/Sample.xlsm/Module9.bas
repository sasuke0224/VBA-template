Attribute VB_Name = "Module9"
Sub mondai8()
    Dim maxVal As Long, minVal As Long

    With Application.WorksheetFunction

    maxVal = .Max(Range(Cells(1, 1), Cells(10,1)))
    minVal = .min(Range(Cells(1, 1), Cells(10,1)))

    End With

    Range("B1") = maxVal
    Range("B2") = minVal

End Sub