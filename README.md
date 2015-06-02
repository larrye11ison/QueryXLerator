# QueryXLerator
Utility for creating Excel (xlsx) files from SQL Queries

## What is this thing?
Have you ever needed to just create a damn Excel file directly from a SQL query batch? This tool does that and then 
gets the hell out of your way.

* Handles multiple resultsets - if you execute a T-SQL batch with multiple queries, each one will be put onto a separate tab.
* Handles SQL datatypes intelligently - applies formatting in Excel appropriately.
* Makes use of Excel Tables...
  * Can format the results using any of the default Excel 2007+ table styles. Even has a semi-non-craptacular UI for making this selection.
  * Can make use of table summary row - you can put several different Excel functions into the summary, like Sum(), Average(), etc.
* Queries execute asyncronously, although they can't currently be canceled... YET.

## What does it look like?
![Main Interface](/WikiAssets/sample.png)

## I wish I could use this awesome library programmatically
You can.

The _QueryXLerator.Library_ assembly can also be used to programmatically write files and/or add data to existing worksheets. This 
can behave like the types of external data linking functionality that's built in to Excel, but without the requirement that the 
external data be in a specific fully-qualified static path. In other words, your data is injected directly into the file with 
your reporting (charts, graphs, etc.).

For instance, you can set up a "template" file with pivot tables, charts, etc. that reference a table with a specific name. 
Then you can use the QXL library to "inject" a new set of datainto the workbook using that same table name. That table's contents
will be overwritten with new data and the next time you open the spreadsheet, everything referencing that data will be updated.
