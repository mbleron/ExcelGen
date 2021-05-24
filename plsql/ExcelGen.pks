create or replace package ExcelGen is
/* ======================================================================================

  MIT License

  Copyright (c) 2020-2021 Marc Bleron

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
====================================================================================== */

  -- file types
  FILE_XLSX       constant pls_integer := 0;
  FILE_XLSB       constant pls_integer := 1;

  -- compatible versions for encryption
  OFFICE2007SP1   constant pls_integer := 0;
  OFFICE2007SP2   constant pls_integer := 1;
  OFFICE2010      constant pls_integer := 2;
  OFFICE2013      constant pls_integer := 3;
  OFFICE2016      constant pls_integer := 4;
  
  /*
  type CT_BorderPr is record (
    style  varchar2(16)
  , color  varchar2(8)
  );
  
  type CT_Border is record (
    left     CT_BorderPr
  , right    CT_BorderPr
  , top      CT_BorderPr
  , bottom   CT_BorderPr
  , content  varchar2(32767)
  );
  
  type CT_Font is record (
    name     varchar2(64)
  , b        boolean := false
  , i        boolean := false
  , color    varchar2(8)
  , sz       pls_integer
  , content  varchar2(32767)
  );

  type CT_PatternFill is record (
    patternType  varchar2(32)
  , fgColor      varchar2(8)
  , bgColor      varchar2(8)
  );

  type CT_Fill is record (
    patternFill  CT_PatternFill
  , content      varchar2(32767)
  );
  
  type CT_CellAlignment is record (
    horizontal  varchar2(16)
  , vertical    varchar2(16)
  , content     varchar2(32767)
  );
  */
  subtype CT_BorderPr is ExcelTypes.CT_BorderPr;
  subtype CT_Border is ExcelTypes.CT_Border;
  subtype CT_Font is ExcelTypes.CT_Font;
  subtype CT_PatternFill is ExcelTypes.CT_PatternFill;
  subtype CT_Fill is ExcelTypes.CT_Fill;
  subtype CT_CellAlignment is ExcelTypes.CT_CellAlignment;
  
  subtype ctxHandle is pls_integer;
  subtype sheetHandle is pls_integer;
  subtype cellStyleHandle is pls_integer;
  
  subtype uint8 is pls_integer range 0..255;

  procedure setDebug (
    p_status in boolean
  );

  function makeRgbColor (
    r  in uint8
  , g  in uint8
  , b  in uint8
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
    p_name   in varchar2
  , p_sz     in pls_integer
  , p_b      in boolean default false
  , p_i      in boolean default false
  , p_color  in varchar2 default null
  )
  return CT_Font;

  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill;

  function makeAlignment (
    p_horizontal  in varchar2 default null
  , p_vertical    in varchar2 default null
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

  function createContext (
    p_type  in pls_integer default FILE_XLSX 
  )
  return ctxHandle;

  procedure closeContext (
    p_ctxId  in ctxHandle 
  );

  procedure addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  );
  
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

  procedure addSheetFromCursor (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  );

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

  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in number
  );

  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  );

  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
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

  -- to be deprecated
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

  -- to be deprecated
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

  procedure setTimestampFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  );

  procedure setEncryption (
    p_ctxId       in ctxHandle
  , p_password    in varchar2
  , p_compatible  in pls_integer default OFFICE2007SP2
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

end ExcelGen;
/
