Option Explicit

' Set the project password.
CreateObject("WScript.Shell").Environment("PROCESS") _
	.Item("APP_DEBUG_PASSWORD") = "tele$ExcelSQLParser"

' Run the main project workbook.
With CreateObject("Excel.Application")
	.Visible = True
	Call .Workbooks.Open( _
		CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) _
			& "\SQLParser.xlsm" _
	)
End With
