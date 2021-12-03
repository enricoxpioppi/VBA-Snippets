Sheets.Add(Before:=Sheets("Input")).Name = "NewSheet"
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "NewSheet"
Sheets.Move(Before:=Sheets("Input")).Name = "NewSheet"