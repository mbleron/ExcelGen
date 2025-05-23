create or replace package ExcelGen is
/* ======================================================================================

  MIT License

  Copyright (c) 2020-2025 Marc Bleron

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.

=========================================================================================
    Change history :
    Marc Bleron       2020-04-01     Creation
    Marc Bleron       2020-05-13     Added CellAlignment style
    Marc Bleron       2020-06-26     Added encryption
    Marc Bleron       2021-04-29     Fixed wrong sheet order in resulting workbook
    Marc Bleron       2021-04-29     Added optional parameter p_sheetIndex in 
                                     addSheetFromXXX routines
    Marc Bleron       2021-05-13     Added XLSB support
    Lee Lindley       2021-07-25     Added setNumFormat
    Marc Bleron       2021-08-22     Added setColumnFormat and setXXXFormat overloads
    Marc Bleron       2021-09-15     Fixed serialization for plain NUMBER values
                                     Fixed invalid control characters in string values
    Marc Bleron       2021-09-04     Added wrapText attribute
    Marc Bleron       2022-02-06     Fixed table format issue for empty dataset
    Marc Bleron       2022-02-15     Added custom column header and width
    Marc Bleron       2022-02-27     Added custom column style
    Marc Bleron       2022-08-23     Fixed number format for column width (xlsx)
    Marc Bleron       2022-09-04     Added row properties
                                     Added multitable sheet model and cell API
                                     Refactoring
    Marc Bleron       2022-11-18     Renamed makeCellRef parameters
    Marc Bleron       2022-11-19     Added gradientFill
    Marc Bleron       2022-11-20     Fixed streamable flag in createWorksheet
    Marc Bleron       2023-02-02     Broken style inheritance between sheet and 
                                     descendant levels
    Lee Lindley,
    Marc Bleron       2023-02-03     Added getRowCount function
    Marc Bleron       2023-02-14     Added p_headerStyle to setTableColumnProperties
    Marc Bleron       2023-02-15     Added Rich Text support
    Marc Bleron       2023-07-26     Added CLOB query string support
    Marc Bleron       2023-07-29     Added Dublin Core properties
    Marc Bleron       2023-09-02     Added p_maxRows to query-related routines
    Marc Bleron       2024-02-23     Added font strikethrough, text rotation, indent
    Marc Bleron       2024-05-10     Added sheet state, formula support
    Marc Bleron       2024-07-21     Added hyperlink, excluded columns, table naming
    Marc Bleron       2024-08-14     Added makeCellRange, data validation
    Marc Bleron       2024-09-06     Conditional formatting
    Marc Bleron       2025-02-08     Image support
    Marc Bleron       2025-05-08     Sheet background
====================================================================================== */

  -- file types
  FILE_XLSX       constant pls_integer := 0;
  FILE_XLSB       constant pls_integer := 1;
  
  -- sheet visibility state
  ST_VISIBLE      constant pls_integer := 0;
  ST_HIDDEN       constant pls_integer := 1;
  ST_VERYHIDDEN   constant pls_integer := 2;

  -- compatible versions for encryption
  OFFICE2007SP1   constant pls_integer := 0;
  OFFICE2007SP2   constant pls_integer := 1;
  OFFICE2010      constant pls_integer := 2;
  OFFICE2013      constant pls_integer := 3;
  OFFICE2016      constant pls_integer := 4;
  
  -- table anchor position
  TOP_LEFT        constant pls_integer := 1;
  TOP_RIGHT       constant pls_integer := 2;
  BOTTOM_RIGHT    constant pls_integer := 3;
  BOTTOM_LEFT     constant pls_integer := 4;
  
  -- drawing anchor type
  TWOCELL_ANCHOR   constant pls_integer := 0;
  ONECELL_ANCHOR   constant pls_integer := 1;
  ABSOLUTE_ANCHOR  constant pls_integer := 2;
  
  -- drawing behaviour when anchor cells are moved or resized
  MOVE_RESIZE          constant pls_integer := 0;  -- twoCell
  MOVE_NO_RESIZE       constant pls_integer := 1;  -- oneCell
  NO_MOVE_NO_RESIZE    constant pls_integer := 2;  -- absolute
  
  subtype CT_BorderPr is ExcelTypes.CT_BorderPr;
  subtype CT_Border is ExcelTypes.CT_Border;
  subtype CT_Font is ExcelTypes.CT_Font;
  subtype CT_GradientStop is ExcelTypes.CT_GradientStop;
  subtype CT_GradientStopList is ExcelTypes.CT_GradientStopList;
  subtype CT_Fill is ExcelTypes.CT_Fill;
  subtype CT_CellAlignment is ExcelTypes.CT_CellAlignment;
  
  subtype ctxHandle is pls_integer;
  subtype sheetHandle is pls_integer;
  subtype cellStyleHandle is pls_integer;
  subtype tableHandle is pls_integer;
  subtype cfmtStyleHandle is pls_integer;
  
  subtype uint8 is ExcelTypes.uint8;

  function getProductName return varchar2;
  
  procedure setDebug (
    p_status in boolean
  );

  function makeRgbColor (
    r  in uint8
  , g  in uint8
  , b  in uint8
  , a  in number default null
  )
  return varchar2;
  
  function makeBorderPr (
    p_style  in varchar2 default null
  , p_color  in varchar2 default null
  )
  return CT_BorderPr;
  
  function makeBorder (
    p_left    in CT_BorderPr default makeBorderPr()
  , p_right   in CT_BorderPr default makeBorderPr()
  , p_top     in CT_BorderPr default makeBorderPr()
  , p_bottom  in CT_BorderPr default makeBorderPr()
  )
  return CT_Border;

  function makeBorder (
    p_style  in varchar2
  , p_color  in varchar2 default null
  )
  return CT_Border;

  function makeFont (
    p_name       in varchar2 default null
  , p_sz         in pls_integer default null
  , p_b          in boolean default null
  , p_i          in boolean default null
  , p_color      in varchar2 default null
  , p_u          in varchar2 default null
  , p_vertAlign  in varchar2 default null
  , p_strike     in boolean default null
  )
  return CT_Font;

  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill;

  function makeGradientStop (
    p_position  in number
  , p_color     in varchar2
  )
  return CT_GradientStop;
  
  function makeGradientFill (
    p_degree  in number default null
  , p_stops   in CT_GradientStopList default null
  )
  return CT_Fill;

  procedure addGradientStop (
    p_fill      in out nocopy CT_Fill
  , p_position  in number
  , p_color     in varchar2
  );

  function makeAlignment (
    p_horizontal    in varchar2 default null
  , p_vertical      in varchar2 default null
  , p_wrapText      in boolean default false
  , p_textRotation  in number default null
  , p_verticalText  in boolean default false
  , p_indent        in number default null
  )
  return CT_CellAlignment;
  
  function makeCellStyle (
    p_ctxId       in ctxHandle
  , p_numFmtCode  in varchar2 default null
  , p_font        in CT_Font default null
  , p_fill        in CT_Fill default null
  , p_border      in CT_Border default null
  , p_alignment   in CT_CellAlignment default null
  )
  return cellStyleHandle;

  function makeCellStyleCss (
    p_ctxId  in ctxHandle
  , p_css    in varchar2
  )
  return cellStyleHandle;

  function makeCondFmtStyle (
    p_ctxId       in ctxHandle
  , p_numFmtCode  in varchar2 default null
  , p_font        in CT_Font default null
  , p_fill        in CT_Fill default null
  , p_border      in CT_Border default null
  , p_alignment   in CT_CellAlignment default null
  )
  return cfmtStyleHandle;

  function makeCondFmtStyleCss (
    p_ctxId  in ctxHandle
  , p_css    in varchar2
  )
  return cfmtStyleHandle;

  function makeCellRef (
    p_colIdx  in pls_integer
  , p_rowIdx  in pls_integer
  )
  return varchar2;

  function makeCellRange (
    p_startRowIdx  in pls_integer
  , p_startColIdx  in pls_integer
  , p_endRowIdx    in pls_integer default null
  , p_endColIdx    in pls_integer default null
  )
  return ExcelTypes.ST_Ref;

  function colPxToCharWidth (p_px in pls_integer) return number;
  function rowPxToPt (p_px in pls_integer) return number;

  function createContext (
    p_type  in pls_integer default FILE_XLSX 
  )
  return ctxHandle;

  procedure closeContext (
    p_ctxId  in ctxHandle 
  );

  function addSheet (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_tabColor    in varchar2 default null
  , p_sheetIndex  in pls_integer default null
  , p_state       in pls_integer default null
  )
  return sheetHandle;

  procedure addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  );

  procedure addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in clob
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  );
  
  function addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle;

  function addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in clob
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle;

  procedure addSheetFromCursor (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  );

  function addSheetFromCursor (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_maxRows     in integer default null
  , p_state       in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle;
  
  function addTable (
    p_ctxId            in ctxHandle
  , p_sheetId          in sheetHandle
  , p_query            in varchar2
  , p_paginate         in boolean default false
  , p_pageSize         in pls_integer default null
  , p_anchorRowOffset  in pls_integer default null
  , p_anchorColOffset  in pls_integer default null
  , p_anchorTableId    in tableHandle default null
  , p_anchorPosition   in pls_integer default null
  , p_maxRows          in integer default null
  , p_excludeCols      in varchar2 default null
  )
  return tableHandle;
  
  function addTable (
    p_ctxId            in ctxHandle
  , p_sheetId          in sheetHandle
  , p_query            in clob
  , p_paginate         in boolean default false
  , p_pageSize         in pls_integer default null
  , p_anchorRowOffset  in pls_integer default null
  , p_anchorColOffset  in pls_integer default null
  , p_anchorTableId    in tableHandle default null
  , p_anchorPosition   in pls_integer default null
  , p_maxRows          in integer default null
  , p_excludeCols      in varchar2 default null
  )
  return tableHandle;

  function addTable (
    p_ctxId            in ctxHandle
  , p_sheetId          in sheetHandle
  , p_rc               in sys_refcursor
  , p_paginate         in boolean default false
  , p_pageSize         in pls_integer default null
  , p_anchorRowOffset  in pls_integer default null
  , p_anchorColOffset  in pls_integer default null
  , p_anchorTableId    in tableHandle default null
  , p_anchorPosition   in pls_integer default null
  , p_maxRows          in integer default null
  , p_excludeCols      in varchar2 default null
  )
  return tableHandle;

  procedure addTableColumn (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in tableHandle
  , p_name      in varchar2
  , p_value     in varchar2
  , p_refStyle  in pls_integer default null
  );
  
  procedure addTableColumnBefore (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in tableHandle
  , p_name      in varchar2
  , p_value     in varchar2
  , p_columnId  in pls_integer
  , p_refStyle  in pls_integer default null
  );

  procedure addTableColumnAfter (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in tableHandle
  , p_name      in varchar2
  , p_value     in varchar2
  , p_columnId  in pls_integer
  , p_refStyle  in pls_integer default null
  );

  procedure addTableHyperlinkColumn (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_tableId         in tableHandle
  , p_name            in varchar2
  , p_location        in varchar2
  , p_linkName        in varchar2 default null
  );

  procedure addTableHyperlinkColumnBefore (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_name       in varchar2
  , p_columnId   in pls_integer
  , p_location   in varchar2
  , p_linkName   in varchar2 default null
  );
  
  procedure addTableHyperlinkColumnAfter (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_name       in varchar2
  , p_columnId   in pls_integer
  , p_location   in varchar2
  , p_linkName   in varchar2 default null
  );

  procedure putDefinedName (
    p_ctxId     in ctxHandle
  , p_name      in varchar2
  , p_value     in varchar2
  , p_scope     in sheetHandle default null
  , p_comment   in varchar2 default null
  , p_cellRef   in varchar2 default null
  , p_refStyle  in pls_integer default null
  );

  procedure putNumberCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in number
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putStringCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in varchar2
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putDateCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in date
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putRichTextCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in varchar2
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putFormulaCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in varchar2
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  , p_refStyle        in pls_integer default null
  );

  procedure putCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in anydata default null
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putHyperlinkCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_location        in varchar2
  , p_linkName        in varchar2 default null
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure putImageCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_image           in blob
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  );

  procedure addImage (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_image       in blob
  , p_anchorType  in pls_integer
  , p_posX        in varchar2 default null
  , p_posY        in varchar2 default null
  , p_extX        in varchar2 default null
  , p_extY        in varchar2 default null
  , p_fromCol     in pls_integer default null
  , p_fromColOff  in varchar2 default null
  , p_fromRow     in pls_integer default null
  , p_fromRowOff  in varchar2 default null
  , p_toCol       in pls_integer default null
  , p_toColOff    in varchar2 default null
  , p_toRow       in pls_integer default null
  , p_toRowOff    in varchar2 default null
  , p_imageProps  in pls_integer default null
  );

  procedure addDataValidationRule (
    p_ctxId             in ctxHandle
  , p_sheetId           in sheetHandle
  , p_type              in varchar2
  , p_cellRange         in ExcelTypes.ST_Sqref
  , p_value1            in varchar2
  , p_value2            in varchar2 default null
  , p_operator          in varchar2 default null
  , p_allowBlank        in boolean default true
  , p_showDropDown      in boolean default null
  , p_showErrorMessage  in boolean default null
  , p_errorMsg          in varchar2 default null
  , p_errorTitle        in varchar2 default null
  , p_errorStyle        in varchar2 default null
  , p_showInputMessage  in boolean default null
  , p_promptMsg         in varchar2 default null
  , p_promptTitle       in varchar2 default null
  , p_refStyle1         in pls_integer default null
  , p_refStyle2         in pls_integer default null
  );

  procedure addCondFormattingRule (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_type        in pls_integer
  , p_cellRange   in ExcelTypes.ST_Sqref
  , p_style       in ExcelGen.cfmtStyleHandle default null
  , p_operator    in pls_integer default null
  , p_value1      in varchar2 default null
  , p_value2      in varchar2 default null
  , p_param       in pls_integer default null
  , p_percent     in boolean default null
  , p_cfvoList    in ExcelTypes.CT_CfvoList default null
  , p_hideValue   in boolean default null
  , p_iconSet     in pls_integer default null
  , p_reverse     in boolean default null
  , p_stopIfTrue  in boolean default null
  , p_refStyle1   in pls_integer default null
  , p_refStyle2   in pls_integer default null
  );

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in number
  );
  
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  );

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in date
  );

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in number
  );

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  );

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in date
  );

  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in number
  );

  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  );

  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in date
  );

  procedure setSheetProperties (
    p_ctxId                in ctxHandle
  , p_sheetId              in sheetHandle
  , p_activePaneAnchorRef  in varchar2 default null
  , p_showGridLines        in boolean default null
  , p_showRowColHeaders    in boolean default null
  , p_defaultRowHeight     in number default null
  , p_background           in blob default null
  );

  procedure mergeCells (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_range    in varchar2
  , p_style    in cellStyleHandle default null
  );

  procedure mergeCells (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowOffset       in pls_integer
  , p_colOffset       in pls_integer
  , p_rowSpan         in pls_integer
  , p_colSpan         in pls_integer
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  , p_style           in cellStyleHandle default null
  );

  procedure setRangeStyle (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_range           in varchar2
  , p_style           in cellStyleHandle
  , p_outsideBorders  in boolean default false
  );

  procedure setRangeStyle (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowOffset       in pls_integer
  , p_colOffset       in pls_integer
  , p_rowSpan         in pls_integer
  , p_colSpan         in pls_integer
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  , p_style           in cellStyleHandle
  , p_outsideBorders  in boolean default false
  );

  procedure setTableProperties (
    p_ctxId              in ctxHandle
  , p_sheetId            in sheetHandle
  , p_tableId            in tableHandle
  , p_style              in varchar2 default null
  , p_showFirstColumn    in boolean default false
  , p_showLastColumn     in boolean default false
  , p_showRowStripes     in boolean default true
  , p_showColumnStripes  in boolean default false
  , p_tableName          in varchar2 default null
  );

  procedure setTableHeader (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_tableId     in tableHandle
  , p_style       in cellStyleHandle default null
  , p_autoFilter  in boolean default false
  );

  procedure setTableColumnProperties (
    p_ctxId        in ctxHandle
  , p_sheetId      in sheetHandle
  , p_tableId      in pls_integer
  , p_columnId     in pls_integer
  , p_columnName   in varchar2 default null
  , p_style        in cellStyleHandle default null
  , p_headerStyle  in cellStyleHandle default null
  );

  procedure setTableColumnFormat (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in pls_integer
  , p_columnId  in pls_integer
  , p_format    in varchar2
  );

  procedure setTableColumnValidationRule (
    p_ctxId             in ctxHandle
  , p_sheetId           in sheetHandle
  , p_tableId           in tableHandle
  , p_columnId          in pls_integer
  , p_type              in varchar2
  , p_value1            in varchar2
  , p_value2            in varchar2 default null
  , p_operator          in varchar2 default null
  , p_allowBlank        in boolean default true
  , p_showDropDown      in boolean default null
  , p_showErrorMessage  in boolean default null
  , p_errorMsg          in varchar2 default null
  , p_errorTitle        in varchar2 default null
  , p_errorStyle        in varchar2 default null
  , p_showInputMessage  in boolean default null
  , p_promptMsg         in varchar2 default null
  , p_promptTitle       in varchar2 default null
  , p_refStyle1         in pls_integer default null
  , p_refStyle2         in pls_integer default null
  );

  procedure setTableRowProperties (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_tableId  in pls_integer
  , p_rowId    in pls_integer default null
  , p_style    in cellStyleHandle default null
  , p_height   in number default null
  );

  procedure addTableCondFmtRule (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_tableId     in tableHandle
  , p_columnId    in pls_integer
  , p_type        in pls_integer
  , p_style       in ExcelGen.cfmtStyleHandle default null
  , p_operator    in pls_integer default null
  , p_value1      in varchar2 default null
  , p_value2      in varchar2 default null
  , p_param       in pls_integer default null
  , p_percent     in boolean default null
  , p_cfvoList    in ExcelTypes.CT_CfvoList default null
  , p_hideValue   in boolean default null
  , p_iconSet     in pls_integer default null
  , p_reverse     in boolean default null
  , p_stopIfTrue  in boolean default null
  , p_refStyle1   in pls_integer default null
  , p_refStyle2   in pls_integer default null
  );

  -- DEPRECATED
  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  );

  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  );

  -- DEPRECATED
  procedure setTableFormat (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_style      in varchar2 default null
  );

  procedure setTableFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_style    in varchar2 default null
  );

  procedure setDateFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  );

  procedure setDateFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  );

  procedure setNumFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  );

  procedure setNumFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  );

  procedure setTimestampFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  );

  procedure setTimestampFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  );

  procedure setColumnFormat (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_columnId  in pls_integer
  , p_format    in varchar2 default null
  , p_header    in varchar2 default null
  , p_width     in number default null
  );
  
  procedure setColumnProperties (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_columnId  in pls_integer
  , p_style     in cellStyleHandle default null
  , p_header    in varchar2 default null
  , p_width     in number default null
  );

  procedure setRowProperties (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_rowId    in pls_integer
  , p_style    in cellStyleHandle default null
  , p_height   in number default null
  );

  procedure setDefaultStyle (
    p_ctxId    in ctxHandle
  , p_style    in cellStyleHandle
  );

  procedure setDefaultStyle (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_style    in cellStyleHandle
  );

$if NOT $$no_crypto OR $$no_crypto IS NULL $then
  procedure setEncryption (
    p_ctxId       in ctxHandle
  , p_password    in varchar2
  , p_compatible  in pls_integer default OFFICE2007SP2
  );
$end

  procedure setCellReferenceStyle (
    p_ctxId     in ctxHandle
  , p_refStyle  in pls_integer
  );

  procedure setCoreProperties (
    p_ctxId        in ctxHandle
  , p_creator      in varchar2 default null
  , p_description  in varchar2 default null
  , p_subject      in varchar2 default null
  , p_title        in varchar2 default null
  );

  function getFileContent (
    p_ctxId  in ctxHandle
  )
  return blob;
  
  procedure createFile (
    p_ctxId      in ctxHandle
  , p_directory  in varchar2
  , p_filename   in varchar2
  );

  function getRowCount (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle 
  , p_tableId  in tableHandle default null
  ) 
  return pls_integer;

end ExcelGen;
/
