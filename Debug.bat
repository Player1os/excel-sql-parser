@ECHO OFF

:: Set the project password.
SET APP_DEBUG_PASSWORD=tele$SQLParser

:: Run the main project workbook.
CALL "%~dp0SQLParser.xlsm"
