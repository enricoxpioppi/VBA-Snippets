'Return file name as String
Application.GetOpenFilename(filefilter, fileindex, title, MultiSelect)

'Open selected file
Set NewWorkBook = Workbooks.Open(Application.GetOpenFilename())
