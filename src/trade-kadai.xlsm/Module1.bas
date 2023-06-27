Attribute VB_Name = "Module1"
Sub kadai2()
    Dim newBookName As String
    Dim newBookPath As String
    Dim newBook As Workbook
    
    newBookName = "output.xlsx"
    
    newBookPath = ThisWorkbook.Path & "\" & newBookName
    
    If Dir(newBookPath) = "" Then

        Set newBook = Workbooks.Add
        
        newBook.SaveAs newBookPath
    
    End If
End Sub
