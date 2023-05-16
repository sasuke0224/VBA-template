Attribute VB_Name = "Module8"
Sub ���7()
    Dim arr(1 To 5) As String
    Dim j As Integer 
    Dim rng As Range
 

    Worksheets("Sheet1").Select
    Range("A1") = "Liam"
    Range("A2") = "Noah"
    Range("A3") = "Elijah"
    Range("A4") = "Oliver"
    Range("A5") = "Lucas"

    arr(1) = Range("A1").Value
    arr(2) = Range("A2").Value
    arr(3) = Range("A3").Value
    arr(4) = Range("A4").Value
    arr(5) = Range("A5").Value

    Randomize

    For Each rng In Selection
     rng.Value = arr(Int((Rnd * 5) + 1))
    Next

    j = 1
    for j = 1 to 5
    Worksheets("Sheet2").Cells(j,1).Value = rng.Value
    j = j + 1
    Next j

End Sub
