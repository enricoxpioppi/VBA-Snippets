Set CBDTemplate = Workbooks.Open(Application.GetOpenFilename(Title:="Please, select CBD template."))
CBDTemplate.SaveAs Filename:="AAAAA", FileFormat:=xlOpenXMLWorkbookMacroEnabled
