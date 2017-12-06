# ![](XLSXWorkbook_38209) XLSX Workbook
**Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel**

<b>Release 28</b> is released to fix a bug introduced by copy-n-paste from Release 27.  Wben I added the code to output a grid to a spreadsheet in the SaveGridtoWorkbook() method, I copied it from the SaveGridtoWorkbookEx() method; however, I failed to initialize a variable.  I read somewhere that most bugs are introduced by copy-n-paste and it would be better to always retype – but I usually take the copy-n-paste route…  Also, a new parameter has been added to these methods to control outputting hidden columns (see documentation).

<b>Release 29</b> is released to add several suggestions by Doug Hennig and to fix reading a spreadsheet with in-line text.

<b>Release 30</b> is released to check for the length of a sheet name to ensure the maximum allowed length is not exceeded; see the release notes for more information.

![](XLSXWorkbook_38236)

Project Manager: [Greg Green](http://www.codeplex.com/site/users/view/gagreen1214)

**![](XLSXWorkbook_38362) [release:Latest release of XLSX Workbook](617822)**
