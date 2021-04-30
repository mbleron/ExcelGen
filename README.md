
# ExcelGen - An Oracle PL/SQL Generator for MS Excel Files

ExcelGen is a PL/SQL utility to create Excel files (.xlsx) out of SQL data sources (query strings or cursors), with automatic pagination over multiple sheets.  
It supports basic formatting options for the header, and table layout.

## Content
* [What's New in...](#whats-new-in)  
* [Bug tracker](#bug-tracker)  
* [Installation](#installation)  
* [Quick Start](#quick-start)  
* [ExcelGen Subprograms and Usage](#excelgen-subprograms-and-usage)  
* [Style specifications](#style-specifications)  
* [Examples](#examples)  
* [CHANGELOG](#changelog)  

## What's New in...

> Version 1.0 : added encryption features  
> Version 0.1b : Beta version

## Bug tracker

Found a bug, have a question, or an enhancement request?  
Please create an issue [here](https://github.com/mbleron/ExcelGen/issues).

## Installation

### Getting source code

Clone this repository or download it as a zip archive.  

Clone or download [MSUtilities](https://github.com/mbleron/MSUtilities).  

### Database requirement

ExcelGen requires Oracle Database 11\.2\.0\.1 and onwards.

### PL/SQL

Using SQL*Plus, connect to the target database schema, then :  
1. Install [MSUtilities](https://github.com/mbleron/MSUtilities) packages.  
2. Install ExcelGen using script [`install.sql`](./install.sql).  

## Quick Start

Basic Excel export from a SQL query :  
```
declare
  ctxId  ExcelGen.ctxHandle;
begin
  ctxId := ExcelGen.createContext();  
  ExcelGen.addSheetFromQuery(ctxId, 'sheet1', 'select * from my_table');
  ExcelGen.setHeader(ctxId, 'sheet1', p_frozen => true);
  ExcelGen.createFile(ctxId, 'TEST_DIR', 'my_file.xlsx');
  ExcelGen.closeContext(ctxId);
end;
/
```

See the following sections for more examples and detailed description of ExcelGen features.

## ExcelGen Subprograms and Usage

* [createContext](#createcontext-function)  
* [closeContext](#closecontext-procedure)  
* [addSheetFromQuery](#addsheetfromquery-procedure-and-function)  
* [addSheetFromCursor](#addsheetfromcursor-procedure-and-function)  
* [setBindVariable](#setbindvariable-procedure)  
* [setHeader](#setheader-procedure)  
* [setTableFormat](#settableformat-procedure)  
* [setDateFormat](#setdateformat-procedure)  
* [setTimestampFormat](#settimestampformat-procedure)  
* [setEncryption](#setencryption-procedure)  
* [getFileContent](#getfilecontent-function)  
* [createFile](#createfile-procedure)  
* [makeRgbColor](#makergbcolor-function)  
* [makeBorderPr](#makeborderpr-function)  
* [makeBorder](#makeborder-function)  
* [makeFont](#makefont-function)  
* [makePatternFill](#makepatternfill-function)  
* [makeAlignment](#makealignment-function)  
* [makeCellStyle](#makecellstyle-function)  
---

### createContext function
This function creates and returns a new generator handle.  

```sql
function createContext
return ctxHandle;
```

---
### closeContext procedure
Releases a context handle previously opened by [createContext](#createcontext-function) function.

```sql
procedure closeContext (
  p_ctxId  in ctxHandle 
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes

---
### addSheetFromQuery procedure and function
Adds a new sheet based on a SQL query string, with optional pagination.  
Available both as a procedure and a function.  
The function returns a sheetHandle value to be used with related subprograms [setHeader](#setheader-procedure), [setBindVariable](#setbindvariable-procedure) and [setTableFormat](#settableformat-procedure).  

```sql
procedure addSheetFromQuery (
  p_ctxId       in ctxHandle
, p_sheetName   in varchar2
, p_query       in varchar2
, p_tabColor    in varchar2 default null
, p_paginate    in boolean default false
, p_pageSize    in pls_integer default null
, p_sheetIndex  in pls_integer default null
);
```
```sql
function addSheetFromQuery (
  p_ctxId       in ctxHandle
, p_sheetName   in varchar2
, p_query       in varchar2
, p_tabColor    in varchar2 default null
, p_paginate    in boolean default false
, p_pageSize    in pls_integer default null
, p_sheetIndex  in pls_integer default null
)
return sheetHandle;
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_sheetName`|Sheet name.|Yes
`p_query`|SQL query string. <br/>Bind variables (if any) can be set via [setBindVariable](#setbindvariable-procedure) procedure.|Yes
`p_tabColor`|Tab [color](#color-specification) of the new sheet.|No
`p_paginate`|Enables pagination of the input data source over multiple sheets. <br/>Use `p_pageSize` parameter to control the maximum number of rows per sheet.|No
`p_pageSize`|Maximum number of rows per sheet, when pagination is enabled. <br/>If set to NULL, Excel sheet limit is used (1,048,576 rows).|No
`p_sheetIndex`|Sheet tab index in the workbook. <br/>If omitted, the sheet is added at the end of the list, after the last existing index.|No

**Notes :**  
When pagination is used, three substitution variables may be used to generate unique names from the input sheet name pattern : 
* PNUM : current page number 
* PSTART : first row number of the current page
* PSTOP : last row number of the current page

For example :  
`sheet${PNUM}` will be expanded to `sheet1`, `sheet2`, `sheet3`, etc.  
`${PSTART}-${PSTOP}` will be expanded to `1-1000`, `1001-2000`, `2001-3000`, etc. assuming a page size of 1000 rows.  

The list of sheet indices specified via `p_sheetIndex` may be sparse.  
For example, if one adds sheet 'A' at index 2, sheet 'B' at index 4 and sheet 'C' at index 1, the resulting workbook will show sheets 'C', 'A' and 'B' in that order.

---
### addSheetFromCursor procedure and function
Adds a new sheet based on a weakly-typed ref cursor, with optional pagination.  
Available both as a procedure and a function.  
The function returns a sheetHandle value to be used with related subprograms [setHeader](#setheader-procedure), [setBindVariable](#setbindvariable-procedure) and [setTableFormat](#settableformat-procedure).  

```sql
procedure addSheetFromCursor (
  p_ctxId       in ctxHandle
, p_sheetName   in varchar2
, p_rc          in sys_refcursor
, p_tabColor    in varchar2 default null
, p_paginate    in boolean default false
, p_pageSize    in pls_integer default null
, p_sheetIndex  in pls_integer default null
);
```
```sql
function addSheetFromCursor (
  p_ctxId       in ctxHandle
, p_sheetName   in varchar2
, p_rc          in sys_refcursor
, p_tabColor    in varchar2 default null
, p_paginate    in boolean default false
, p_pageSize    in pls_integer default null
, p_sheetIndex  in pls_integer default null
)
return sheetHandle;
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|Yes
`p_sheetName`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|Yes
`p_rc`|Input ref cursor.|Yes
`p_tabColor`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|No
`p_paginate`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|No
`p_pageSize`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|No
`p_sheetIndex`|Cf. [addSheetFromQuery](#addsheetfromquery-procedure-and-function).|No

---
### setBindVariable procedure
This procedure binds a value to a variable in the SQL query associated with a given sheet, specified by either a sheet name or a sheet handle as returned from `addSheetFromXXX` functions.  
It is overloaded to accept a NUMBER, VARCHAR2 or DATE value.  

```sql
procedure setBindVariable (
  p_ctxId      in ctxHandle
, p_sheetName  in varchar2
, p_bindName   in varchar2
, p_bindValue  in number
);
```
```sql
procedure setBindVariable (
  p_ctxId      in ctxHandle
, p_sheetName  in varchar2
, p_bindName   in varchar2
, p_bindValue  in varchar2
);
```
```sql
procedure setBindVariable (
  p_ctxId      in ctxHandle
, p_sheetName  in varchar2
, p_bindName   in varchar2
, p_bindValue  in date
);
```

```sql
procedure setBindVariable (
  p_ctxId       in ctxHandle
, p_sheetId     in sheetHandle
, p_bindName    in varchar2
, p_bindValue   in number
);
```
```sql
procedure setBindVariable (
  p_ctxId      in ctxHandle
, p_sheetId    in sheetHandle
, p_bindName   in varchar2
, p_bindValue  in varchar2
);
```
```sql
procedure setBindVariable (
  p_ctxId      in ctxHandle
, p_sheetId    in sheetHandle
, p_bindName   in varchar2
, p_bindValue  in date
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_sheetName`|Sheet name.|Yes
`p_sheetId`|Sheet handle.|Yes
`p_bindName`|Bind variable name.|Yes
`p_bindValue`|Bind variable value.|Yes

---  
### setHeader procedure
This procedure adds a header row to a given sheet.  
Column names are derived from the SQL source query.  

```sql
procedure setHeader (
  p_ctxId       in ctxHandle
, p_sheetName   in varchar2
, p_style       in cellStyleHandle default null
, p_frozen      in boolean default false
, p_autoFilter  in boolean default false
);
```

```sql
procedure setHeader (
  p_ctxId       in ctxHandle
, p_sheetId     in sheetHandle
, p_style       in cellStyleHandle default null
, p_frozen      in boolean default false
, p_autoFilter  in boolean default false
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_sheetName`|Sheet name.|Yes
`p_sheetId`|Sheet handle.|Yes
`p_style`|Handle to a cell style created via [makeCellStyle](#makecellstyle-function) function.|No
`p_frozen`|Set this parameter to true in order to freeze the header row.|No
`p_autoFilter`|Set this parameter to true in order to add an automatic filter to this sheet.|No

---
### setTableFormat procedure
This procedure applies a table layout to a given sheet.  

```sql
procedure setTableFormat (
  p_ctxId      in ctxHandle
, p_sheetName  in varchar2
, p_style      in varchar2 default null
);
```

```sql
procedure setTableFormat (
  p_ctxId      in ctxHandle
, p_sheetId    in sheetHandle
, p_style      in varchar2 default null
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_sheetName`|Sheet name.|Yes
`p_sheetId`|Sheet handle.|Yes
`p_styleName`|Name of a predefined Excel table style to apply. <br/>See [Predefined table styles](#predefined-table-styles) for a list of available styles.|No

---
### setDateFormat procedure
This procedure sets the format applied to DATE values in the resulting spreadsheet file.  
The format must follow MS Excel proprietary [syntax](https://support.office.com/en-us/article/format-a-date-the-way-you-want-8e10019e-d5d8-47a1-ba95-db95123d273e).  
Default is `dd/mm/yyyy hh:mm:ss`.  

```sql
procedure setDateFormat (
  p_ctxId   in ctxHandle
, p_format  in varchar2
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_format`|Date format string.|Yes

---
### setTimestampFormat procedure
This procedure sets the format applied to TIMESTAMP values in the resulting spreadsheet file.  
Default is `dd/mm/yyyy hh:mm:ss.000`.  

```sql
procedure setTimestampFormat (
  p_ctxId   in ctxHandle
, p_format  in varchar2
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_format`|Timestamp format string.|Yes


---
### setEncryption procedure
This procedure sets the password used to encrypt the document, along with the minimum compatible Office version necessary to open it.

```sql
procedure setEncryption (
  p_ctxId       in ctxHandle
, p_password    in varchar2
, p_compatible  in pls_integer default OFFICE2007SP2
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_password`|Password.|Yes
`p_compatible`|Minimum compatible Office version for encryption. <br/>One of `OFFICE2007SP1`, `OFFICE2007SP2`, `OFFICE2010`, `OFFICE2013`, `OFFICE2016`. Default is `OFFICE2007SP2`.|No


---
### getFileContent function
This function builds the spreadsheet file and returns it as a temporary BLOB.  

```sql
function getFileContent (
  p_ctxId  in ctxHandle
)
return blob; 
```

---
### createFile procedure
This procedure builds the spreadsheet file and write it directly to a directory.  

```sql
procedure createFile (
  p_ctxId      in ctxHandle
, p_directory  in varchar2
, p_filename   in varchar2
);
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_directory`|Target directory.<br/> Must be a valid Oracle directory name.|Yes
`p_filename`|File name.|Yes

---
### STYLING

### makeRgbColor function
This function builds an RGB hex triplet from individual Red, Green and Blue components supplied as unsigned 8-bit integers (0-255).
For example, `makeRgbColor(219,112,147)` will return `#DB7093`.

```sql
function makeRgbColor (
  r  in uint8
, g  in uint8
, b  in uint8
)
return varchar2;
```

Parameter|Description|Mandatory
---|---|---
`r`|Red component value.|Yes
`g`|Green component value.|Yes
`b`|Blue component value.|Yes

---
### makeBorderPr function
This function builds an instance of a cell border edge, from a border style name and a color.

```sql
function makeBorderPr (
  p_style  in varchar2 default null
, p_color  in varchar2 default null
)
return CT_BorderPr;
```

Parameter|Description|Mandatory
---|---|---
`p_style`|Border style name. <br/>See [Border styles](#border-styles) for a list of available styles.|No
`p_color`|Border [color](#color-specification).|No

---
### makeBorder function
This function builds an instance of a cell border (all edges).  
It is overloaded to accept either individual edge formatting (left, right, top and bottom), or the same format for all edges.

Overload 1 :  
```sql
function makeBorder (
  p_left    in CT_BorderPr default makeBorderPr()
, p_right   in CT_BorderPr default makeBorderPr()
, p_top     in CT_BorderPr default makeBorderPr()
, p_bottom  in CT_BorderPr default makeBorderPr()
)
return CT_Border;
```

Overload 2 :  
```sql
function makeBorder (
  p_style  in varchar2
, p_color  in varchar2 default null
)
return CT_Border;
```

Overload 2 is a shorthand for : 
```
makeBorder(
  p_left    => makeBorderPr(p_style, p_color)
, p_right   => makeBorderPr(p_style, p_color)
, p_top     => makeBorderPr(p_style, p_color)
, p_bottom  => makeBorderPr(p_style, p_color)
)
```

Parameter|Description|Mandatory
---|---|---
`p_left`|Left edge format, as returned by [makeBorderPr](#makeborderpr-function) function.|No
`p_right`|Right edge format, as returned by [makeBorderPr](#makeborderpr-function) function.|No
`p_top`|Top edge format, as returned by [makeBorderPr](#makeborderpr-function) function.|No
`p_bottom`|Bottom edge format, as returned by [makeBorderPr](#makeborderpr-function) function.|No
`p_style`|Border style name.|Yes
`p_color`|Border [color](#color-specification).|No

---
### makeFont function
This function builds an instance of a cell font.

```sql
function makeFont (
  p_name   in varchar2
, p_sz     in pls_integer
, p_b      in boolean default false
, p_i      in boolean default false
, p_color  in varchar2 default null
)
return CT_Font;
```

Parameter|Description|Mandatory
---|---|---
`p_name`|Font name.|Yes
`p_sz`|Font size, in points.|Yes
`p_b`|Bold font style (true\|false).|No
`p_i`|Italic font style (true\|false).|No
`p_color`|Font [color](#color-specification).|No

---
### makePatternFill function
This function builds an instance of a cell pattern fill.

```sql
function makePatternFill (
  p_patternType  in varchar2
, p_fgColor      in varchar2 default null
, p_bgColor      in varchar2 default null
)
return CT_Fill;
```

Parameter|Description|Mandatory
---|---|---
`p_patternType`|Pattern type. <br/>See [Pattern types](#pattern-types) for a list of available types.|Yes
`p_fgColor`|Foreground [color](#color-specification) of the pattern.|No
`p_bgColor`|Background [color](#color-specification) of the pattern.|No

Note :
For a solid fill (no pattern), the color must be specified using the foreground color parameter.  

---
### makeAlignment function
This function builds an instance of a cell alignment.

```sql
function makeAlignment (
  p_horizontal  in varchar2 default null
, p_vertical    in varchar2 default null
)
return CT_CellAlignment;
```

Parameter|Description|Mandatory
---|---|---
`p_horizontal`|Horizontal alignment type, one of `left`, `center` or `right`.|No
`p_vertical`|Vertical alignment type, one of `top`, `center` or `bottom`.|No

---
### makeCellStyle function
This function builds an instance of a cell style, composed of optional number format, font, fill and border specifications, and returns a handle to it.

```sql
function makeCellStyle (
  p_ctxId       in ctxHandle
, p_numFmtCode  in varchar2 default null
, p_font        in CT_Font default null
, p_fill        in CT_Fill default null
, p_border      in CT_Border default null
, p_alignment   in CT_CellAlignment default null
)
return cellStyleHandle;
```

Parameter|Description|Mandatory
---|---|---
`p_ctxId`|Context handle.|Yes
`p_numFmtCode`|Number format code.|No
`p_font`|Font style instance, as returned by [makeFont](#makefont-function) function.|No
`p_fill`|Fill style instance, as returned by [makePatternFill](#makepatternfill-function) function.|No
`p_border`|Border style instance, as returned by [makeBorder](#makeborder-function) function.|No
`p_alignment`|Cell alignment instance, as returned by [makeAlignment](#makealignment-function) function.|No

Example : 

This sample code creates a cell style composed of the following facets : 
* Font : Calibri, 11 pts, bold face
* Fill : YellowGreen solid fill
* Border : Thick red edges
* Alignment : Horizontally centered
```
declare

  ctxId      ExcelGen.ctxHandle;
  cellStyle  ExcelGen.cellStyleHandle;

begin
  ...
 
  cellStyle := ExcelGen.makeCellStyle(
                 p_ctxId     => ctxId
               , p_font      => ExcelGen.makeFont('Calibri',11,true)
               , p_fill      => ExcelGen.makePatternFill('solid','YellowGreen')
               , p_border    => ExcelGen.makeBorder('thick','red')
               , p_alignment => ExcelGen.makeAlignment(horizontal => 'center')
               );
  ...
```

## Style specifications

### Color specification
Color values expected in style-related subprograms may be specified using one of the following conventions : 

* RGB hex triplet, prefixed with a hash sign, e.g. `#7FFFD4`.  
Function [makeRgbColor](#makergbcolor-function) can be used to build such a color code from individual RGB components.

* Named color from the [CSS4 specification](https://www.w3.org/TR/css-color-4/#named-colors), e.g. `Aquamarine`.  

### Predefined table styles  

__Light__  
![TableStyleLight](./resources/tablestylelight.png "TableStyleLight")

__Medium__  
![TableStyleMedium](./resources/tablestylemedium.png "TableStyleMedium")

__Dark__  
![TableStyleDark](./resources/tablestyledark.png "TableStyleDark")

### Border styles

![BorderStyles](./resources/borderstyles.png "BorderStyles")

### Pattern types

![PatternTypes](./resources/patterntypes.png "PatternTypes")

## Examples

* Single query to sheet mapping, with header formatting : [employees.xlsx](./samples/employees.xlsx)  

```
declare

  sqlQuery   varchar2(32767) := 'select * from hr.employees';
  sheetName  varchar2(31 char) := 'Sheet1';
  ctxId      ExcelGen.ctxHandle;
  
begin
  
  ctxId := ExcelGen.createContext();  
  ExcelGen.addSheetFromQuery(ctxId, sheetName, sqlQuery);
      
  ExcelGen.setHeader(
    ctxId
  , sheetName
  , p_style => ExcelGen.makeCellStyle(
                 p_ctxId => ctxId
               , p_font  => ExcelGen.makeFont('Calibri',11,true)
               , p_fill  => ExcelGen.makePatternFill('solid','LightGray')
               )
  , p_frozen     => true
  , p_autoFilter => true
  );
  
  ExcelGen.setDateFormat(ctxId, 'dd/mm/yyyy');
  
  ExcelGen.createFile(ctxId, 'TEST_DIR', 'employees.xlsx');
  ExcelGen.closeContext(ctxId);
  
end;
/
```

* Multiple queries, with table layout : [dept_emp.xlsx](./samples/dept_emp.xlsx)  

```
declare
  ctxId      ExcelGen.ctxHandle;
begin
  ctxId := ExcelGen.createContext();
  
  -- add dept sheet
  ExcelGen.addSheetFromQuery(ctxId, 'dept', 'select * from hr.departments');
  ExcelGen.setHeader(ctxId, 'dept', p_autoFilter => true);
  ExcelGen.setTableFormat(ctxId, 'dept', 'TableStyleLight2');

  -- add emp sheet
  ExcelGen.addSheetFromQuery(ctxId, 'emp', 'select * from hr.employees where salary >= :1 order by salary desc');
  ExcelGen.setBindVariable(ctxId, 'emp', '1', 7000);  
  ExcelGen.setHeader(ctxId, 'emp', p_autoFilter => true);
  ExcelGen.setTableFormat(ctxId, 'emp', 'TableStyleLight7');
  
  ExcelGen.setDateFormat(ctxId, 'dd/mm/yyyy');
  
  ExcelGen.createFile(ctxId, 'TEST_DIR', 'dept_emp.xlsx');
  ExcelGen.closeContext(ctxId);
  
end;
/
```

* Ref cursor paginated over multiple sheets : [all_objects.xlsx](./samples/all_objects.xlsx)

```
declare
  rc         sys_refcursor;
  sheetName  varchar2(128) := 'sheet${PNUM}';
  ctxId      ExcelGen.ctxHandle;
begin
  
  open rc for 
  select * from all_objects where owner = 'SYS';

  ctxId := ExcelGen.createContext();
  
  ExcelGen.addSheetFromCursor(
    p_ctxId     => ctxId
  , p_sheetName => sheetName
  , p_rc        => rc
  , p_tabColor  => 'DeepPink'
  , p_paginate  => true
  , p_pageSize  => 10000
  );
    
  ExcelGen.setHeader(
    ctxId
  , sheetName
  , p_style  => ExcelGen.makeCellStyle(ctxId, p_fill => ExcelGen.makePatternFill('solid','LightGray'))
  , p_frozen => true
  );
  
  ExcelGen.createFile(ctxId, 'TEST_DIR', 'all_objects.xlsx');
  ExcelGen.closeContext(ctxId);
  
end;
/
```  

* Yet another example : [sample4.xlsx](./samples/sample4.xlsx)

```
declare
  ctxId    ExcelGen.ctxHandle;
  sheet1   ExcelGen.sheetHandle;
  sheet2   ExcelGen.sheetHandle;
  sheet3   ExcelGen.sheetHandle;
begin
  
  ctxId := ExcelGen.createContext();
  
  -- adding a new sheet in position 3
  sheet1 := ExcelGen.addSheetFromQuery(ctxId, 'c', 'select * from hr.employees where department_id = :1', p_sheetIndex => 3);
  ExcelGen.setBindVariable(ctxId, sheet1, '1', 30);
  ExcelGen.setTableFormat(ctxId, sheet1, 'TableStyleLight1');
  ExcelGen.setHeader(ctxId, sheet1, p_autoFilter => true, p_frozen => true);
  
  -- adding a new sheet in last position (4)
  sheet2 := ExcelGen.addSheetFromQuery(ctxId, 'b', 'select * from hr.employees');
  ExcelGen.setTableFormat(ctxId, sheet2, 'TableStyleLight2');
  ExcelGen.setHeader(ctxId, sheet2, p_autoFilter => true, p_frozen => true);
  
  -- adding a new sheet in position 1, with a 10-row pagination
  sheet3 := ExcelGen.addSheetFromQuery(ctxId, 'a${PNUM}', 'select * from hr.employees', p_paginate => true, p_pageSize => 10, p_sheetIndex => 1);
  ExcelGen.setHeader(ctxId, sheet3, p_autoFilter => true, p_frozen => true);

  ExcelGen.createFile(ctxId, 'TEST_DIR', 'sample4.xlsx');
  ExcelGen.closeContext(ctxId);
  
end;
/
```  


## CHANGELOG

### 1.1 (2021-04-30)
* Fix : issue #1

### 1.0 (2020-06-28)
* Added encryption

### 0.1b (2020-03-25)
* Beta version


## Copyright and license

Copyright 2020-2021 Marc Bleron. Released under MIT license.
