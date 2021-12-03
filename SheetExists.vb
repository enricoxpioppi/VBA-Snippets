Function SheetExists(sheetToFind As String) As Boolean
    SheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
End Function