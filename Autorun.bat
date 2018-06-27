@ECHO OFF

:: Switch to the deploy directory.
PUSHD \\kedata\Data1\B2B_Business_Inteligence\Osama Hassanein\ExcelSQLParser

:: Set the autorun parameters.
SET APP_IS_AUTORUN_MODE=TRUE
SET APP_INPUT_WORKBOOK_FILE_PATH=H:\INPUT_WORKBOOK.xlsx
SET APP_QUERY_FILE_PATH=H:\QUERY.sql
SET APP_OUTPUT_WORKBOOK_FILE_PATH=H:\OUTPUT_WORKBOOK.xlsx

:: Run the main project workbook.
CALL "SQLParser.xlsm"

:: Return to the original working directory.
POPD
