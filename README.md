# ![](XLSXWorkbook_38209) XLSX Workbook
**Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel**

Release 18 - Fixed some bugs and added new methods.  This release also attempts to reduce the time to save a XLSX file.  Restructuring of the internal cursors and changes to cell formatting were performed to reduce the time to create a new XLSX file.  The cell formatting process has been changed to now define a style for a cell to be formatted with.  These methods are listed in the Release Notes and more explained in the Documentation.  The older cell formatting methods are not removed for backward compatibility; however, these methods are now depreciated and will not be further enhanced.

Release 19 - Fixed bugs in two methods referring to the event name; see Release Notes for more information.  I did not update the Documentation (see the event Method in the class for additional documentation).

I have begun to expand the read capability of this class.  Currently if an xlsx document contains images, text blocks or other content, these objects are lost when the workbook is re-saved.  I am working on the ability to retain these objects; however, I am not currently working on the ability to add new objects (just retain what is in an existing xlsx document).  I decided to release now the code in its current state with the additions for retaining objects removed.  This is to allow for the new functionality added to be used and tested by the community. 


![](XLSXWorkbook_38236)

Project Manager: [Greg Green](http://www.codeplex.com/site/users/view/gagreen1214)

**![](XLSXWorkbook_38362) [release:Latest release of XLSX Workbook](617822)**
