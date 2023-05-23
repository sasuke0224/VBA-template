Attribute VB_Name = "Module10"
Sub mondai9()
    Dim startDate As Date, endDate As Date
    startDate = Sheet1.Range("A1").Value
    endDate = Sheet1.Range("A2").Value

    Dim dayCount As Long
    dayCount = endDate - startDate
    
    Sheet1.Range("A3").Value = dayCount
End Sub
