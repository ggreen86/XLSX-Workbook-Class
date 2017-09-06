# ![](XLSXWorkbook_38209) XLSX Workbook
**Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel**

Release 26 is released to fix a bug in the GetCellValue() method and in the SaveGridToWorkbookEx() / SaveTableToWorkbookEx() methods.
	
In the GetCellValue() method when a cell value was formatted in the workbook as a numeric format with decimal places, but the value was actually entered without any decimal values or decimal point, the method would cause an exception on the CAST() statement due to incorrect data type.

The SaveGridToWorkbookEx() / SaveTableToWorkbookEx() methods were not correctly handling the decimal separator and decimal point values for these settings when not set to a ‘,’ and ‘.’ respectively (this affects the European countries).  I am now saving these settings, changing to USA standard settings and then restoring back to user defaults.
 


![](XLSXWorkbook_38236)

Project Manager: [Greg Green](http://www.codeplex.com/site/users/view/gagreen1214)

**![](XLSXWorkbook_38362) [release:Latest release of XLSX Workbook](617822)**
