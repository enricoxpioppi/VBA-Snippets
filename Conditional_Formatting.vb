Cells.FormatConditions.Delete

Dim Rng2 As Range
Set Rng2 = ActiveSheet.Range(Cells(pdFirstItemRow, pdFirstItemColumn - 1), Cells(pdLastTableRow, pdLastColumn))

Rng2.FormatConditions.Add Type:=xlExpression, Formula1:= _
  "=OR(LEFT($A15,LEN($G$10))=$G$10,LEFT($A15,LEN($H$10))=$H$10,LEFT($A15,LEN($I$10))=$I$10,LEFT($A15,LEN($J$10))=$J$10,LEFT($A15,LEN($K$10))=$K$10,LEFT($A15,LEN($L$10))=$L$10,LEFT($A15,LEN($M$10))=$M$10)"

Rng2.FormatConditions(1).Interior.Color = RGB(91, 155, 213)
