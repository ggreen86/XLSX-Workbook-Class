# XLSX Workbook

Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel

## Release 29

Released to add several suggestions by Doug Hennig and to fix reading a spreadsheet with in-line text.
## Release 30

Released to check for the length of a sheet name to ensure the maximum allowed length is not exceeded; see the release notes for more information.

## Release 31

Released to correct some identified bugs and for a new method; see the release notes for more information.

## Release 32 Beta

This release is a major change to the Read of the workbook and support of more workbook objects (images).  There are a number of fixes for the support of exporting a table and grid directly to a XLSX workbook.  There are a number of new methods for supporting more functionality -- See the Release notes.  Since this is a major change to this class, I am now releasing it as Beta software.  There has been a number of users besides myself that has been testing this class and I have been fixing bugs or incorporating the fixes provided into the class.

## Release 32

Release 32 is now being released as production and no longer BETA.  It has been several months now without any issues being reported.

Released 24 Feb 2021.

## Release 33 Beta

Two corrections were made to the class that affects the methods SaveGridToWorkbook() and SaveGridToWorkbookEx().  The first change is the class's handling of the grid column properties for Format and InputMask.  I have tried to retain the original intent of the formatting when exporting to Excel.  So please report an issue if you find any problems when the grid column formatting is exported.  The second change involves illegal characters that are not permitted in the spreadsheet tab title which I had been checking for (the code converts illegal characters to underscore).  However, if the tab title string passed had these illegal characters encoded as an xml equilvalent then it would cause an error on trying to open the spreadsheet.  I have added additional checks to remove illegal characters if they are encoded.

Additional correction/enhancement made to ensure the temp directory name for the workbook create is unique and not used (suggestion by Doug Hennig).  Another bug was identified in the method SetCellBorderEx() and is now corrected.

## Written By

Gregory Green
