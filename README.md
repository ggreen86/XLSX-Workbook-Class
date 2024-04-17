# XLSX Workbook

Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel

## Release 40

Minor update to allow the GetCellValue() method to retrieve the cell value for a cell that has a formula.  The formula must be first calculated by another spreadsheet program such as MS Excel; the method for adding a formula to a cell does not calculate the cell value.

## Release 41

Release 41 is released to correct defects reported by users in the issues log.  No new functionality has been introduced.  However, some existing methods have been enhanced to either correct behavior, add additional capability, or faster performance.  See the Release Notes for detailed changes.  Now released as production as no issues have been received and I have not experienced any problems.

## Release 42

Release 42 is released to add a new method for direct output a table to xlxs file with columns/rows transposed.  Also, a bug is fixed that caused a loss of formatting for rotated text, vertical & horizontal text cell alignment, cell text indent, and cell text word-wrapping.  See Release Notes for more information.  Released with no further beta testing to production.

I also added a new property, TemporaryPath, that defaults to =SYS(2023).  You can set this another file location of your choice.  All the class temporary files will then be written to that location instead of the system TEMP directory.  I missed adding this to the documentation.

## Written By

Gregory Green
