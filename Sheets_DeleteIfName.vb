If SheetExists("Backlog WIP") = True And SheetExists("SFBA") = True Then
    Application.DisplayAlerts = False
    For Each ws In Application.ActiveWorkbook.Worksheets
        If ws.Name <> "Backlog WIP" And ws.Name <> "SFBA" Then
            ws.Delete
        End If
    Next
End If