# XLSX Workbook

Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel

## Release 34

Bug fixes have been applied; see the Release Notes for explanations.  Also, the reading of an existing spreadsheet has been extensively reworked to increase performance.

## Release 35

Bug fix for an obscure bug in the freeze rows/columns; see the Release Notes for explanations.  Added several methods for handling the Print Area.

## Release 36

Bug fix when reading an existing workbook with cell values formatted as text.  Corrected the read and write processing.

## Release 37

Bug fix for reading workbooks that do not have a defined styles.xml file included in the workbook.  See Release Notes for detailed description of fix.

## Release 38

Bug fix for cell format codes using the date-time #DEFINE values and locale code.  With this release, the documentation has also been updated to include all methods in the class.  Note that the bookmarks will navigate to the page; to see the method you might have to scroll down.

## Release 39

Release to add new methods for Conditional Formatting, new methods for Table Formatting, and fixed the SetColumnBestFit() method to auto set the column width based on the cell contents.  

This release is currently in beta for feedback on implementation.  See the Release Notes and the Documentation on the new methods.  

Beta 3 version fixes a bug in the Reading of the styles.xml that caused the numeric formatting to be lost.  

Beta 4 fixed the SetColumnBestFit() method to now auto set the column widths based on the widest cell entry.  

Beta 5 version fixes a bug if the table name had a space in it (spaces not allowed).  Beta 5 also adds functionality to the SaveGridToWorkbookEx() method to defining a table format to be applied (new parameter).  

Beta 6 removed a check for the file existing on the file system before adding as a Hyperlink.  This check prevented a file link to a URL to be added as a hyperlink and is not needed.

Beta 7 added the CodePage setting to the xl_sheets cursor to account for the name of the sheet tab.

## Written By

Gregory Green
