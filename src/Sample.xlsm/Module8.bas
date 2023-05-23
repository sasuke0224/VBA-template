Attribute VB_Name = "Module8"
Sub ���7()
    Dim arr(1 To 5) As String
    Dim rng As Range
 

    Worksheets("Sheet1").Select
    Range("A1") = "Liam"
    Range("A2") = "Noah"
    Range("A3") = "Elijah"
    Range("A4") = "Oliver"
    Range("A5") = "Lucas"

    arr(1) = Range("A1")
    arr(2) = Range("A2")
    arr(3) = Range("A3")
    arr(4) = Range("A4")
    arr(5) = Range("A5")

    Randomize

    Set sheet = Worksheets("Sheet2")
    Set selectedRange = sheet.Range("A1:A5") 

    For Each rng In selectedRange
     rng.Value = arr(Int((Rnd * 5) + 1))
    Next


End Sub
