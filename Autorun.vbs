' Set the autorun parameters.
With CreateObject("WScript.Shell").Environment("PROCESS")
	.Item("APP_IS_AUTORUN_MODE") = "TRUE"
	.Item("APP_INPUT_WORKBOOK_FILE_PATH") = "H:\INPUT_WORKBOOK.xlsx"
	.Item("APP_QUERY_FILE_PATH") = "H:\QUERY.sql"
	.Item("APP_OUTPUT_WORKBOOK_FILE_PATH") = "H:\OUTPUT_WORKBOOK.xlsx"
End With

' Run the main project workbook.
CreateObject("Excel.Application") _
	.Workbooks.Open(CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\SQLParser.xlsm")
