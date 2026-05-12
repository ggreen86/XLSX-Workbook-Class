# XLSX Workbook

Read from and write to XLSX format files without any automation or export with support for full cell formatting in Excel

## Release 50

Release 50 is released to correct a bug that was discovered when assigning table format to a cell range.  The table name failed the name check when being opened by Excel and caused a recover error.  This is now corrected.  Please see the Release Notes for more information.

## Release 51

Release 51 is released to add a new property that is used in the calculation of column best fit for columns containing numeric values.  See the Release Notes for more information.

## Release 52

Release 52 is released to fix some bugs and compatibility issues with Openpyxl.  Changes were made to streamline when processing the cell values when opening a spreadsheet and how the CodePage property is implemented.  See the Release Notes for more information.

## Release 53

Release 53 is released to fix a bug affecting the cell alignment for grids using dynamic properties.  See the Release Notes for more information.

## Release 54

Release 54 is released to fix a number of bugs.  See the Release Notes for more information.

## Release 55

Release 55 - not released.

## Release 56

Release 56 - **RELEASED IN ALPHA**.  This is being released to allow users to review the changes made for the InsertRow and InsertColumn methods.  The InsertCell method **has not yet been updated**.  These methods for inserting was not accounting for all the possible impacts due to features being added after these methods were initially written.  I am still working on the impact to formulas for inserting rows/columns.  If there are formulas that reference columns or rows that are affected by an insert, this is **currently** not being updated.  I have partial code written that will parse the formulas into the individual tokens.  The current implementation of the InsertRow and InsertColumn should update everything except for formulas.

There are other updates/corrections -- see the included Release Notes.

## Written By

Gregory Green
