# ![](XLSXWorkbook_38209) XLSX Workbook
**Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel**

<b>Release 26</b> is released to fix a bug in the GetCellValue() method and in the SaveGridToWorkbookEx() / SaveTableToWorkbookEx() methods.
	
In the GetCellValue() method when a cell value was formatted in the workbook as a numeric format with decimal places, but the value was actually entered without any decimal values or decimal point, the method would cause an exception on the CAST() statement due to incorrect data type.

The SaveGridToWorkbookEx() / SaveTableToWorkbookEx() methods were not correctly handling the decimal separator and decimal point values for these settings when not set to a ‘,’ and ‘.’ respectively (this affects the European countries).  I am now saving these settings, changing to USA standard settings and then restoring back to user defaults.

<b>Release 27</b> is released to fix a serious bug in the SaveGridToWorkbookEx() method.  If this method was called with the freeze top row parameter set to False, then the spreadsheet would be unreadable when trying to open.  This is fixed.  Also, a suggestion was made to save the column order based on the grid display order which is also incorporated in this release.
 
<b>Release 28</b> is released to fix a bug introduced by copy-n-paste from Release 27.  Wben I added the code to output a grid to a spreadsheet in the SaveGridtoWorkbook() method, I copied it from the SaveGridtoWorkbookEx() method; however, I failed to initialize a variable.  I read somewhere that most bugs are introduced by copy-n-paste and it would be better to always retype – but I usually take the copy-n-paste route…  Also, a new parameter has been added to these methods to control outputting hidden columns (see documentation).

![](XLSXWorkbook_38236)

Project Manager: [Greg Green](http://www.codeplex.com/site/users/view/gagreen1214)

**![](XLSXWorkbook_38362) [release:Latest release of XLSX Workbook](617822)**
