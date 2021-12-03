'**//  Declaration                                            '

  Dim pSheet As Worksheet, dSheet As Worksheet
  Dim pTable As PivotTable
  Dim pCache As PivotCache
  Dim pRange As Range
  Dim pLastRow As Long, pLastColumn As Long

'**\\                                                                 '
'**//  Define sheets                                        '

  Set pSheet = Worksheets("Summary")
  Set dSheet = Worksheets("Data")

'**\\                                                                 '
'**//  Define data range                                 '

With dSheet
    pLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    pLastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    Set pRange = .Range("A1").Resize(pLastRow, pLastCol) ' set data range for Pivot Table
End With

'**\\                                                                 '
'**//  Pivot cache                                            '

Set pCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pRange, Version:=xlPivotTableVersion14)

'**\\                                                                 '
'**//  First time running the macro?              '

On Error Resume Next
Set pTable = pSheet.PivotTables("Summary") 'check if "ERA_Dashboard" Pivot Table already created (in past runs of this Macro)

On Error GoTo 0
If pTable Is Nothing Then
    'create a new Pivot Table in "Summary" sheet
        Set pTable = pSheet.PivotTables.Add(PivotCache:=pCache, TableDestination:=pSheet.Range("A1"), TableName:="Summary")
Else
    'refresh the Pivot cache with the updated Range
        pTable.ChangePivotCache pCache
        pTable.RefreshTable
End If

Worksheets("Summary").Activate

'**\\                                                                 '
'**//  Add fields                                               '

  '**//**//  Segments                                        '

    With Worksheets("Summary").PivotTables("Summary").PivotFields("Segment")
      .Orientation = xlRowField
      .Position = 1
    End With

  '**\\**\\                                             '
  '**//**//  Revenues                                        '

    For i = uCurrentYear To LastYear
      Dim pName As String
      pName = "" & "Sum of " & i & ""
      Dim ptField As String
      ptField = "" & i & ""
      pSheet.PivotTables("Summary").AddDataField pSheet.PivotTables("Summary").PivotFields(ptField), pName, xlSum
    Next i

  '**\\**\\                                                           '

'**\\                                                                 '
