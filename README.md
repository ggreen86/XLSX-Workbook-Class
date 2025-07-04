# XLSX Workbook

Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel

## Release 40

Minor update to allow the GetCellValue() method to retrieve the cell value for a cell that has a formula.  The formula must be first calculated by another spreadsheet program such as MS Excel; the method for adding a formula to a cell does not calculate the cell value.

## Release 41

Release 41 is released to correct defects reported by users in the issues log.  No new functionality has been introduced.  However, some existing methods have been enhanced to either correct behavior, add additional capability, or faster performance.  See the Release Notes for detailed changes.  Now released as production as no issues have been received and I have not experienced any problems.

## Release 42

Release 42 is released to add a new method for direct output a table to xlxs file with columns/rows transposed.  Also, a bug is fixed that caused a loss of formatting for rotated text, vertical & horizontal text cell alignment, cell text indent, and cell text word-wrapping.  See Release Notes for more information.  Released with no further beta testing to production.

I also added a new property, TemporaryPath, that defaults to =SYS(2023).  You can set this another file location of your choice.  All the class temporary files will then be written to that location instead of the system TEMP directory.  I missed adding this to the documentation.

## Release 43

Release 43 is released to add methods for read-only access to a workbook (very fast access), fix some bugs, and speed improvements to opening workbooks to the method OpenXlsxWorkbook().  The documentation has been updated with the new methods added; however, the bookmarks have not yet been added for quick access.  Please see the Release Notes and full documentation for more information.

## Release 44

Release 44 is released to correct an oversight on my part in the new method OpenXlsxWorkbookEx().  I missed the processing of cells with in-line string values and cells with formulas.  Also, I missed converting text strings from the xml encoded value to the normal value.  There is also change in the parameter line for the method OpenXlsxWorkbookEx().  Please see the Release Notes and full documentation for more information.

## Release 45

Release 45 is released.  This version includes some bug fixes that were reported by users via the Issues, adds enhancements to some existing methods, and adds new methods for additional features.  The additional features includes the ability to define column formulas when exporting from grids and being able to apply cell formatting based on built-in standard Styles.   Please see the Release Notes and full documentation for more information.

## Release 46 - Release removed due to insufficient regression testing

Release 46 is released.  This version adds enhancements to existing methods; please see the Release Notes for details.  

## Release 47

Release 47 is released.  Corrected logic error in setting the cell merging calculation for the report listener.

## Release 48

Release 48 is released for an enhancement to the method SaveTableToWorkbookEx() in order to allow for creating spreadsheets from tables with a row count that exceeds the sheet row limit.  See Release Notes for more information.

## Release 49

Release 49 is released to provide a new method for adding a sheet to an existing workbook file; see method SaveTableToWorkbookSheetEx().  Also, I have added a "fudge" factor for when calculating the best-fit for columns containing numeric values in an attempt to address two reported issues.  Additionally, the documentation has been updated for the new method and some typos corrected.  See Release Notes for more information.

## Release 50

Release 50 is released to correct a bug that was discovered when assigning table format to a cell range.  The table name failed the name check when being opened by Excel and caused a recover error.  This is now corrected.  Please see the Release Notes for more information.

## Written By

Gregory Green
