create or replace package body ExcelGen is

  NAMED_COLORS          constant varchar2(4000) := 
  'aliceblue:F0F8FF;antiquewhite:FAEBD7;aqua:00FFFF;aquamarine:7FFFD4;azure:F0FFFF;beige:F5F5DC;bisque:FFE4C4;black:000000;' ||
  'blanchedalmond:FFEBCD;blue:0000FF;blueviolet:8A2BE2;brown:A52A2A;burlywood:DEB887;cadetblue:5F9EA0;chartreuse:7FFF00;chocolate:D2691E;' ||
  'coral:FF7F50;cornflowerblue:6495ED;cornsilk:FFF8DC;crimson:DC143C;cyan:00FFFF;darkblue:00008B;darkcyan:008B8B;darkgoldenrod:B8860B;' ||
  'darkgray:A9A9A9;darkgreen:006400;darkgrey:A9A9A9;darkkhaki:BDB76B;darkmagenta:8B008B;darkolivegreen:556B2F;darkorange:FF8C00;darkorchid:9932CC;' ||
  'darkred:8B0000;darksalmon:E9967A;darkseagreen:8FBC8F;darkslateblue:483D8B;darkslategray:2F4F4F;darkslategrey:2F4F4F;darkturquoise:00CED1;darkviolet:9400D3;' ||
  'deeppink:FF1493;deepskyblue:00BFFF;dimgray:696969;dimgrey:696969;dodgerblue:1E90FF;firebrick:B22222;floralwhite:FFFAF0;forestgreen:228B22;' ||
  'fuchsia:FF00FF;gainsboro:DCDCDC;ghostwhite:F8F8FF;gold:FFD700;goldenrod:DAA520;gray:808080;green:008000;greenyellow:ADFF2F;' ||
  'grey:808080;honeydew:F0FFF0;hotpink:FF69B4;indianred:CD5C5C;indigo:4B0082;ivory:FFFFF0;khaki:F0E68C;lavender:E6E6FA;' ||
  'lavenderblush:FFF0F5;lawngreen:7CFC00;lemonchiffon:FFFACD;lightblue:ADD8E6;lightcoral:F08080;lightcyan:E0FFFF;lightgoldenrodyellow:FAFAD2;lightgray:D3D3D3;' ||
  'lightgreen:90EE90;lightgrey:D3D3D3;lightpink:FFB6C1;lightsalmon:FFA07A;lightseagreen:20B2AA;lightskyblue:87CEFA;lightslategray:778899;lightslategrey:778899;' ||
  'lightsteelblue:B0C4DE;lightyellow:FFFFE0;lime:00FF00;limegreen:32CD32;linen:FAF0E6;magenta:FF00FF;maroon:800000;mediumaquamarine:66CDAA;' ||
  'mediumblue:0000CD;mediumorchid:BA55D3;mediumpurple:9370DB;mediumseagreen:3CB371;mediumslateblue:7B68EE;mediumspringgreen:00FA9A;mediumturquoise:48D1CC;mediumvioletred:C71585;' ||
  'midnightblue:191970;mintcream:F5FFFA;mistyrose:FFE4E1;moccasin:FFE4B5;navajowhite:FFDEAD;navy:000080;oldlace:FDF5E6;olive:808000;' ||
  'olivedrab:6B8E23;orange:FFA500;orangered:FF4500;orchid:DA70D6;palegoldenrod:EEE8AA;palegreen:98FB98;paleturquoise:AFEEEE;palevioletred:DB7093;' ||
  'papayawhip:FFEFD5;peachpuff:FFDAB9;peru:CD853F;pink:FFC0CB;plum:DDA0DD;powderblue:B0E0E6;purple:800080;rebeccapurple:663399;' ||
  'red:FF0000;rosybrown:BC8F8F;royalblue:4169E1;saddlebrown:8B4513;salmon:FA8072;sandybrown:F4A460;seagreen:2E8B57;seashell:FFF5EE;' ||
  'sienna:A0522D;silver:C0C0C0;skyblue:87CEEB;slateblue:6A5ACD;slategray:708090;slategrey:708090;snow:FFFAFA;springgreen:00FF7F;' ||
  'tan:D2B48C;teal:008080;thistle:D8BFD8;tomato:FF6347;turquoise:40E0D0;violet:EE82EE;wheat:F5DEB3;white:FFFFFF;' ||
  'steelblue:4682B4;whitesmoke:F5F5F5;yellow:FFFF00;yellowgreen:9ACD32';

  -- OPC part MIME types
  MT_STYLES          constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  MT_WORKBOOK        constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  MT_WORKSHEET       constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  MT_SHAREDSTRINGS   constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
  MT_TABLE           constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
  --MT_COMMENTS        constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
  
  -- Binary MIME types
  MT_STYLES_BIN         constant varchar2(256) := 'application/vnd.ms-excel.styles';
  MT_WORKSHEET_BIN      constant varchar2(256) := 'application/vnd.ms-excel.worksheet';
  MT_SHAREDSTRINGS_BIN  constant varchar2(256) := 'application/vnd.ms-excel.sharedStrings';
  MT_TABLE_BIN          constant varchar2(256) := 'application/vnd.ms-excel.table';
  
  -- Relationship types
  RS_OFFICEDOCUMENT  constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RS_WORKSHEET       constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  RS_STYLES          constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  RS_SHAREDSTRINGS   constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
  RS_TABLE           constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
  --RS_COMMENTS        constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';

  RANGE_EMPTY_REF        constant varchar2(100) := 'Range error : empty reference';
  RANGE_INVALID_REF      constant varchar2(100) := 'Range error : invalid reference ''%s''';
  RANGE_INVALID_COL      constant varchar2(100) := 'Range error : column out of range ''%s''';
  RANGE_INVALID_ROW      constant varchar2(100) := 'Range error : row out of range ''%d''';
  RANGE_INVALID_EXPR     constant varchar2(100) := 'Range error : invalid range expression ''%s''';
  RANGE_START_ROW_ERR    constant varchar2(100) := 'Range error : start row (%d) must be lower or equal than end row (%d)';
  RANGE_START_COL_ERR    constant varchar2(100) := 'Range error : start column (''%s'') must be lower or equal than end column (''%s'')';
  RANGE_EMPTY_COL_REF    constant varchar2(100) := 'Range error : missing column reference in ''%s''';
  RANGE_EMPTY_ROW_REF    constant varchar2(100) := 'Range error : missing row reference in ''%s'''; 

  MAX_COLUMN_NUMBER      constant pls_integer := 16384;
  MAX_ROW_NUMBER         constant pls_integer := 1048576;
  MAX_BUFFER_SIZE        constant pls_integer := 32767;

  DIGITS                 constant varchar2(10) := '0123456789';
  LETTERS                constant varchar2(26) := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  
  CNTRL_CHARS            constant varchar2(32) := to_char(unistr('\0000\0001\0002\0003\0004\0005\0006\0007\0008\000B\000C\000E\000F\0010\0011\0012\0013\0014\0015\0016\0017\0018\0019\001A\001B\001C\001D\001E\001F'));
  
  DEFAULT_DATE_FMT       constant varchar2(32) := 'dd/mm/yyyy hh:mm:ss';
  DEFAULT_TIMESTAMP_FMT  constant varchar2(32) := 'dd/mm/yyyy hh:mm:ss.000';
  DEFAULT_NUM_FMT        constant varchar2(32) := null; 
  NLS_PARAM_STRING       constant varchar2(32) := 'nls_numeric_characters=''. ''';
  
  -- supertypes
  ST_NUMBER              constant pls_integer := 0;
  ST_STRING              constant pls_integer := 1;
  ST_DATETIME            constant pls_integer := 2;
  ST_LOB                 constant pls_integer := 3;

  buffer_too_small       exception;
  pragma exception_init (buffer_too_small, -19011);

  type color_map_t is table of varchar2(6) index by varchar2(20);

  type stream_t is record (
    content   clob
  , buf       varchar2(32767)
  , buf_sz    pls_integer
  );

  type data_t is record (
    varchar2_value  varchar2(32767)
  , char_value      char(32767)
  , number_value    number
  , date_value      date
  , ts_value        timestamp
  , tstz_value      timestamp with time zone
  , clob_value      clob
  );
  
  type data_map_t is table of data_t index by pls_integer;
  type data_row_t is record (dataMap data_map_t);
  
  type cell_ref_t is record (value varchar2(10), c varchar2(3), cn pls_integer, r pls_integer); 
  type range_t is record (expr varchar2(32), start_ref cell_ref_t, end_ref cell_ref_t);
  
  --type int32_list_t is table of pls_integer;
  type int_set_t is table of pls_integer index by pls_integer;
  type link_token_map_t is table of varchar2(8) index by pls_integer;
  type link_t is record (target varchar2(2048), tooltip varchar2(256), tokens link_token_map_t, fmla varchar2(8192));
  --type link_rel_map_t is table of varchar2(256) index by varchar2(2048);
  type link_map_t is table of varchar2(2048) index by pls_integer;

  type column_ref_list_t is table of varchar2(3);
  
  type column_t is record (
    name     varchar2(128)
  , type     pls_integer
  , scale    pls_integer
  , id       pls_integer
  , colRef   varchar2(3)
  , xfId     pls_integer := 0
  , hasLink  boolean := false
  , link     link_t
  , supertype pls_integer
  , excluded  boolean := false
  );
    
  type column_list_t is table of column_t;
  type column_map_t is table of pls_integer index by varchar2(128);
  type column_set_t is table of varchar2(3) index by pls_integer;

  type string_map_t is table of pls_integer index by varchar2(32767);
  type string_list_t is table of varchar2(32767);
  
  type CT_Relationship is record (
    Type    varchar2(256)
  , Id      varchar2(256)
  , Target  varchar2(2048)
  );
  
  type CT_Relationships is table of CT_Relationship;
  
  type part_t is record (
    name         varchar2(256)
  , contentType  varchar2(256)
  , content      clob
  , contentBin   blob
  , rels         CT_Relationships
  , isBinary     boolean := false
  );
  
  type part_list_t is table of part_t;
  type part_index_map_t is table of pls_integer index by varchar2(256);

  type zip_entry_t is record (
    offset    integer
  , filename  varchar2(256)
  , content   blob
  );
  
  type zip_entry_list_t is table of zip_entry_t;
  
  type zip_t is record (
    content  blob
  , entries  zip_entry_list_t
  );
  
  type package_t is record (
    parts        part_list_t
  , partIndices  part_index_map_t
  , rels         CT_Relationships
  , content      blob
  );
  
  type defaultFmts_t is record (
    dateFmt       varchar2(128)
  , timestampFmt  varchar2(128)
  , numFmt        varchar2(128)
  );
  
  type CT_BorderMap is table of pls_integer index by varchar2(32767);
  type CT_Borders is table of CT_Border index by pls_integer;
  
  type CT_FontMap is table of pls_integer index by varchar2(32767);
  type CT_Fonts is table of CT_Font index by pls_integer;
  
  type CT_FillMap is table of pls_integer index by varchar2(32767);
  type CT_Fills is table of CT_Fill index by pls_integer;
  
  type CT_NumFmtMap is table of pls_integer index by varchar2(32767);
  type CT_NumFmts is table of varchar2(32767) index by pls_integer;
  
  type CT_Xf is record (
    numFmtId   pls_integer := 0
  , fontId     pls_integer := 0
  , fillId     pls_integer := 0
  , borderId   pls_integer := 0
  , xfId       pls_integer
  , alignment  CT_CellAlignment := null
  , content    varchar2(32767)
  );
  
  type CT_CellXfMap is table of pls_integer index by varchar2(32767);
  type CT_CellXfs is table of CT_Xf index by pls_integer;
  
  type CT_CellStyle is record (
    name       varchar2(256)
  , xfId       pls_integer
  , builtinId  pls_integer
  );
  
  type CT_CellStyles is table of CT_CellStyle;
  
  type CT_Stylesheet is record (
    numFmtMap       CT_NumFmtMap
  , numFmts         CT_NumFmts
  , fontMap         CT_FontMap
  , fonts           CT_Fonts
  , fillMap         CT_FillMap
  , fills           CT_Fills
  , borderMap       CT_BorderMap
  , borders         CT_Borders
  , cellStyleXfMap  CT_CellXfMap
  , cellStyleXfs    CT_CellXfs
  , cellXfMap       CT_CellXfMap
  , cellXfs         CT_CellXfs
  , cellStyles      CT_CellStyles
  , hasHlink        boolean := false
  , hlinkXfId       pls_integer
  );
  
  type CT_TableColumn is record (
    id    pls_integer
  , name  varchar2(256)
  );
  
  type CT_TableColumns is table of CT_TableColumn;
  
  type CT_Table is record (
    id          pls_integer
  , name        varchar2(256)
  , ref         range_t
  , cols        CT_TableColumns
  , showHeader  boolean
  , autoFilter  boolean
  , styleName   varchar2(64)
  , partName    varchar2(256)
  );
  
  type CT_Tables is table of CT_Table index by pls_integer;
  
  type CT_TableParts is table of pls_integer;
  
  type CT_Sheet is record (
    name         varchar2(128)
  , sheetId      pls_integer
  , rId          varchar2(256)
  , partName     varchar2(256)
  , filterRange  range_t
  , filterXti    pls_integer
  , tableParts   CT_TableParts
  );
  
  type CT_Sheets is table of CT_Sheet;
  type CT_SheetMap is table of pls_integer index by varchar2(128);
  
  type CT_Workbook is record (
    sheets           CT_Sheets
  , sheetMap         CT_SheetMap
  , styles           CT_Stylesheet
  , tables           CT_Tables
  , hasDefinedNames  boolean := false
  );
  
  type bind_variable_t is record (
    name   varchar2(30)
  , value  anydata
  );
  
  type bind_variable_list_t is table of bind_variable_t;
  
  type sql_metadata_t is record (
    queryString       varchar2(32767)
  , cursorNumber      integer
  , bindVariables     bind_variable_list_t
  , columnList        column_list_t
  , columnMap         column_map_t
  , visibleColumnSet  column_set_t
  , excludeSet        int_set_t
  , partitionBySize   boolean := false
  , partitionSize     pls_integer
  , partitionId       pls_integer
  , r_num             pls_integer
  );
  
  type sheet_header_t is record (
    show        boolean
  , xfId        pls_integer
  , isFrozen    boolean
  , autoFilter  boolean
  );
  
  type sheet_column_t is record (
    header  varchar2(1024)
  , fmt     varchar2(128)
  , width   number
  );
  
  type sheet_column_map_t is table of sheet_column_t index by pls_integer;
  --type column_fmt_map_t is table of varchar2(128) index by pls_integer;
  
  type sheet_definition_t is record (
    sheetName      varchar2(128)
  , sheetIndex     pls_integer
  , tabColor       varchar2(8)
  , header         sheet_header_t
  , formatAsTable  boolean
  , tableStyle     varchar2(32)
  , sqlMetadata    sql_metadata_t
  , defaultFmts    defaultFmts_t
  --, columnFmtMap   column_fmt_map_t
  , columnMap      sheet_column_map_t
  , columnLinkMap  link_map_t
  );
  
  type sheet_definition_map_t is table of sheet_definition_t index by pls_integer;
  
  type encryption_info_t is record (
    version       varchar2(3)
  , cipherName    varchar2(16)
  , hashFuncName  varchar2(16)
  , password      varchar2(512)
  );
  
  type context_t is record (
    string_map           string_map_t
  , string_list          string_list_t := string_list_t()
  , string_cnt           pls_integer := 0
  , column_ref_list      column_ref_list_t := column_ref_list_t()
  , workbook             CT_Workbook
  , pck                  package_t
  , sheetDefinitionMap   sheet_definition_map_t
  , sheetIndexMap        CT_SheetMap
  , defaultFmts          defaultFmts_t
  , encryptionInfo       encryption_info_t
  , fileType             pls_integer
  );
  
  type context_cache_t is table of context_t index by pls_integer;
  
  colorMap       color_map_t;
  ctx_cache      context_cache_t;
  debug_enabled  boolean := false;
  
  procedure setDebug (p_status in boolean)
  is
  begin
    debug_enabled := p_status;
  end;

  procedure debug (message in varchar2)
  is
  begin
    if debug_enabled then
      dbms_output.put_line(to_char(systimestamp, 'HH24:MI:SS.FF3')||' '||message);
    end if;
  end;

  procedure error (
    message in varchar2
  , arg1    in varchar2 default null
  , arg2    in varchar2 default null
  , arg3    in varchar2 default null
  , code    in number default -20800
  ) 
  is
  begin
    raise_application_error(code, utl_lms.format_message(message, arg1, arg2, arg3));
  end;
  
  procedure initColorMap 
  is
    token  varchar2(32);
    p1     pls_integer := 1;
    p2     pls_integer;  
    i      pls_integer;
  begin
    debug('initColorMap');
    loop
      p2 := instr(NAMED_COLORS, ';', p1);
      if p2 = 0 then
        token := substr(NAMED_COLORS, p1);
      else
        token := substr(NAMED_COLORS, p1, p2-p1);    
        p1 := p2 + 1;
      end if;
      i := instr(token,':');
      colorMap(substr(token,1,i-1)) := substr(token,i+1);
      exit when p2 = 0;
    end loop;   
  end;
  
  procedure init
  is  
  begin
    initColorMap;
  end;

  function base26decode (p_str in varchar2) 
  return pls_integer 
  result_cache
  is
    l_result  pls_integer;
    l_base    pls_integer := 1;
  begin
    if p_str is not null then
      l_result := 0;
      for i in 1 .. length(p_str) loop
        l_result := l_result + (ascii(substr(p_str,-i,1)) - 64) * l_base;
        l_base := l_base * 26;
      end loop;
    end if;
    return l_result;
  end;

  function base26encode (p_num in pls_integer) 
  return varchar2
  is
    l_result  varchar2(3);
    l_num     pls_integer := p_num;
  begin
    if p_num is not null then
      while l_num != 0 loop
        l_result := chr(65 + mod(l_num-1,26)) || l_result;
        l_num := trunc((l_num-1)/26);
      end loop;
    end if;
    return l_result;
  end;
  
  function escapeQuote (str in varchar2)
  return varchar2 is
  begin
    return replace(str,'"','""');
  end;

  function enquote (str in varchar2) 
  return varchar2 is
  begin
    return '"'||escapeQuote(str)||'"';
  end;

  function stripXmlControlChars (str in varchar2)
  return varchar2
  is
  begin
    return translate(str, '_'||CNTRL_CHARS, '_');
  end;  

  function escapeXmlControlChars (str in varchar2)
  return varchar2
  is
    output  varchar2(32767);
  begin
    -- using a bunch of replace's instead of a loop
    output := replace(str, chr(0), '_x0000_');
    output := replace(output, chr(1), '_x0001_');
    output := replace(output, chr(2), '_x0002_');
    output := replace(output, chr(3), '_x0003_');
    output := replace(output, chr(4), '_x0004_');
    output := replace(output, chr(5), '_x0005_');
    output := replace(output, chr(6), '_x0006_');
    output := replace(output, chr(7), '_x0007_');
    output := replace(output, chr(8), '_x0008_');
    output := replace(output, chr(11), '_x000B_');
    output := replace(output, chr(12), '_x000C_');
    output := replace(output, chr(14), '_x000E_');
    output := replace(output, chr(15), '_x000F_');
    output := replace(output, chr(16), '_x0010_');
    output := replace(output, chr(17), '_x0011_');
    output := replace(output, chr(18), '_x0012_');
    output := replace(output, chr(19), '_x0013_');
    output := replace(output, chr(20), '_x0014_');
    output := replace(output, chr(21), '_x0015_');
    output := replace(output, chr(22), '_x0016_');
    output := replace(output, chr(23), '_x0017_');
    output := replace(output, chr(24), '_x0018_');
    output := replace(output, chr(25), '_x0019_');
    output := replace(output, chr(26), '_x001A_');
    output := replace(output, chr(27), '_x001B_');
    output := replace(output, chr(28), '_x001C_');
    output := replace(output, chr(29), '_x001D_');
    output := replace(output, chr(30), '_x001E_');
    output := replace(output, chr(31), '_x001F_');
    return output;
  end;

  function int2raw (int32 in binary_integer, sz in pls_integer default null) return raw
  is
    r raw(4) := utl_raw.cast_from_binary_integer(int32, utl_raw.little_endian);
  begin
    return case when sz is not null then utl_raw.substr(r, 1, sz) else r end;
  end;

  function parseIntList (input in varchar2, sep in varchar2)
  return int_set_t
  is
    i       pls_integer;
    token   varchar2(256);
    p1      pls_integer := 1;
    p2      pls_integer;
    output  int_set_t;
  begin
    if input is not null then
      loop
        p2 := instr(input, sep, p1);
        if p2 = 0 then
          token := substr(input, p1);
        else
          token := substr(input, p1, p2-p1);    
          p1 := p2 + 1;
        end if;
        begin
          i := to_number(trim(token));
          if i is not null then
            output(i) := i;
          end if;
        exception
          when value_error then
            error('Invalid numeric token ''%s''', token);
        end;
        exit when p2 = 0;
      end loop;
    end if;
    return output;
  end;
  
  function makeCellRef (
    colRef  in varchar2
  , rowRef  in pls_integer
  )
  return cell_ref_t
  is
    cellRef cell_ref_t;
  begin
    cellRef.c := colRef;
    cellRef.cn := base26decode(cellRef.c);
    cellRef.r := rowRef;
    cellRef.value := cellRef.c || to_char(cellRef.r);
    return cellRef;
  end;
  
  function makeRange (
    startCol  in varchar2
  , startRow  in pls_integer
  , endCol    in varchar2
  , endRow    in pls_integer
  )
  return range_t
  is
    range  range_t;
  begin
    range.start_ref := makeCellRef(startCol, startRow);
    range.end_ref := makeCellRef(endCol, endRow);
    range.expr := range.start_ref.value || case when range.end_ref.value is not null then ':'||range.end_ref.value end;
    return range;
  end;
  
  function parseRangeExpr (
    expr  in varchar2 
  )
  return range_t
  is
    pos    pls_integer;
    range  range_t;
    
    procedure readCellRef (expr in varchar2, cellRef in out nocopy cell_ref_t) is
      colRef  varchar2(32);
      rowRef  varchar2(32);
      colNum  pls_integer;
      rowNum  pls_integer;
    begin
      if expr is null then
        error(RANGE_EMPTY_REF);
      end if;
      colRef := rtrim(expr, DIGITS);
      rowRef := ltrim(expr, LETTERS);
      if rtrim(rowRef, DIGITS) is not null or rtrim(colRef, LETTERS) is not null then
        error(RANGE_INVALID_REF, expr);
      end if;
      colNum := base26decode(colRef);
      -- validate column reference
      if colNum > MAX_COLUMN_NUMBER then
        error(RANGE_INVALID_COL, colRef);
      end if;
      rowNum := to_number(rowRef);
      if rowNum not between 1 and MAX_ROW_NUMBER then
        error(RANGE_INVALID_ROW, rowNum);
      end if;
      cellRef.r := rowNum;
      cellRef.c := colRef; 
      cellRef.cn := colNum;
    end;
    
  begin
    
    if expr is not null then
      
      pos := instr(expr, ':');
      if pos != 0 then
        readCellRef(substr(expr, 1, pos-1), range.start_ref);
        readCellRef(substr(expr, pos+1), range.end_ref);
        -- validate range :
        if range.start_ref.c is not null and range.end_ref.c is null 
          or range.start_ref.c is null and range.end_ref.c is not null 
          or range.start_ref.r is not null and range.end_ref.r is null 
          or range.start_ref.r is null and range.end_ref.r is not null
        then
          error(RANGE_INVALID_EXPR, expr);
        elsif range.start_ref.r > range.end_ref.r then
          error(RANGE_START_ROW_ERR, range.start_ref.r, range.end_ref.r);
        elsif range.start_ref.cn > range.end_ref.cn then
          error(RANGE_START_COL_ERR, range.start_ref.c, range.end_ref.c);
        end if;
                
      else
        readCellRef(expr, range.start_ref);
        -- validate single cell reference
        if range.start_ref.c is null then
          error(RANGE_EMPTY_COL_REF, expr);
        elsif range.start_ref.r is null then
          error(RANGE_EMPTY_ROW_REF, expr);
        end if;
      end if;
    
    end if;
    
    return range;
    
  end;
  
  function getRangeExpr (
    range     in range_t
  , anchored  in boolean default false
  )
  return varchar2
  is
    function getAnchoredRefValue (cellRef in cell_ref_t) 
    return varchar2
    is
    begin
      return '$' || cellRef.c || '$' || to_char(cellRef.r);
    end;
  begin
    return case when anchored then
             getAnchoredRefValue(range.start_ref) || 
             case when range.end_ref.value is not null then ':'||getAnchoredRefValue(range.end_ref) end
           else
             range.expr
           end;
  end;

  procedure writeBlobToFile (
    p_directory  in varchar2
  , p_filename   in varchar2
  , p_content    in blob
  )
  is
    MAX_BUF_SIZE  constant pls_integer := 32767;
    file       utl_file.file_type;
    pos        integer := 1;
    chunkSize  pls_integer := dbms_lob.getchunksize(p_content);
    amt        pls_integer := least(trunc(MAX_BUF_SIZE/chunkSize)*chunkSize, MAX_BUF_SIZE);
    buf        raw(32767);
  begin
    file := utl_file.fopen(p_directory, p_filename, 'wb', 32767);
    loop
      begin
        dbms_lob.read(p_content, amt, pos, buf);
      exception
        when no_data_found then
          exit;
      end;
      utl_file.put_raw(file, buf);
      pos := pos + amt;
    end loop;
    utl_file.fclose(file);
  end;

  function xmlToBlob (
    input       in clob
  , encoding    in varchar2 default 'UTF-8'
  , standalone  in boolean default true
  )
  return blob
  is
    dest_offset    integer;
    src_offset     integer := 1;
    charset_id     pls_integer := nls_charset_id(utl_i18n.map_charset(encoding, flag => utl_i18n.IANA_TO_ORACLE));
    lang_context   integer := dbms_lob.default_lang_ctx;
    warning        integer;
    output         blob;
    xmlProlog      raw(256) := utl_raw.cast_to_raw(
                                 '<?xml version="1.0" encoding="'||encoding||'"'||
                                 case when standalone is not null 
                                      then ' standalone="'||case when standalone then 'yes' else 'no' end||'"' 
                                 end ||
                                 '?>'
                               );
    xmlPrologSize  pls_integer := utl_raw.length(xmlProlog);
  begin
    dbms_lob.createtemporary(output, true);
    dbms_lob.writeappend(output, xmlPrologSize, xmlProlog);
    dest_offset := xmlPrologSize + 1;
    dbms_lob.convertToBlob(
      dest_lob     => output
    , src_clob     => input
    , amount       => dbms_lob.getlength(input)
    , dest_offset  => dest_offset
    , src_offset   => src_offset
    , blob_csid    => charset_id
    , lang_context => lang_context
    , warning      => warning
    );
    return output;
  end;

  procedure string_write (
    buf  in out nocopy varchar2
  , str  in varchar2
  )
  is
  begin
    buf := buf || str;
  end;
  
  function validateColor (
    colorCode  in varchar2
  )
  return varchar2
  is
    rgbCode  varchar2(8);
  begin
    if colorCode is not null then
      -- RGB color code?
      if substr(colorCode,1,1) = '#' then
        rgbCode := upper(substr(colorCode,2)); 
        if rgbCode is null or not regexp_like(rgbCode, '^[0-9A-F]{6}$') then
          error('Invalid RGB color code: %s', colorCode);
        end if;
        -- adding opaque alpha channel by default
        rgbCode := 'FF' || rgbCode;
      elsif colorMap.exists(lower(colorCode)) then
        rgbCode := 'FF' || colorMap(lower(colorCode));
      elsif regexp_like(colorCode,'^theme:\d+$') then
        rgbCode := colorCode;
      else
        error('Invalid color code: %s', colorCode);
      end if;
    end if;
    return rgbCode;
  end;
  
  function makeRgbColor (
    r  in uint8
  , g  in uint8
  , b  in uint8
  )
  return varchar2
  is
  begin
    return '#' || to_char(r * 65536 + g * 256 + b, 'FM0XXXXX');
  end;

  function putNumfmt (
    styles  in out nocopy CT_Stylesheet
  , fmt     in varchar2 
  )
  return pls_integer
  is
    numFmtId  pls_integer;
  begin
    if styles.numFmtMap.exists(fmt) then
      numFmtId := styles.numFmtMap(fmt);
    else
      numFmtId := nvl(styles.numFmts.last, 163) + 1;
      styles.numFmts(numFmtId) := fmt;
      styles.numFmtMap(fmt) := numFmtId;
    end if;
    return numFmtId;
  end;
  
  function makeBorderPr (
    p_style  in varchar2 default null
  , p_color  in varchar2 default null
  )
  return CT_BorderPr
  is
    borderPr  CT_BorderPr;
  begin
    borderPr.style := p_style;
    borderPr.color := validateColor(p_color);
    return borderPr;
  end;
  
  procedure setBorderContent (
    border  in out nocopy CT_Border
  )
  is
    function getBorderPrContent (borderName in varchar2, borderPr in CT_BorderPr)
    return varchar2
    is
    begin
      return '<' || borderName || 
             case when borderPr.style is not null then ' style="'||borderPr.style||'"' end || 
             case when borderPr.color is not null then '><color rgb="'||borderPr.color||'"/></'||borderName||'>' else '/>' end;
    end;
  begin
    string_write(border.content, '<border>');
    string_write(border.content, getBorderPrContent('left', border.left));
    string_write(border.content, getBorderPrContent('right', border.right));
    string_write(border.content, getBorderPrContent('top', border.top));
    string_write(border.content, getBorderPrContent('bottom', border.bottom));
    string_write(border.content, '</border>');    
  end;
  
  function makeBorder (
    p_left    in CT_BorderPr default makeBorderPr()
  , p_right   in CT_BorderPr default makeBorderPr()
  , p_top     in CT_BorderPr default makeBorderPr()
  , p_bottom  in CT_BorderPr default makeBorderPr()
  )
  return CT_Border
  is
    border  CT_Border;
  begin
    border.left := p_left;
    border.right := p_right;
    border.top := p_top;
    border.bottom := p_bottom;
    setBorderContent(border);
    return border;
  end;
  
  function makeBorder (
    p_style  in varchar2
  , p_color  in varchar2 default null
  )
  return CT_Border
  is
    borderPr  CT_BorderPr := makeBorderPr(p_style, p_color);
  begin
    return makeBorder(borderPr, borderPr, borderPr, borderPr);
  end;
  
  function putBorder (
    styles  in out nocopy CT_Stylesheet
  , border  in CT_Border
  )
  return pls_integer
  is
    borderId  pls_integer;
  begin
    if styles.borderMap.exists(border.content) then
      borderId := styles.borderMap(border.content);
    else
      borderId := nvl(styles.borders.last, -1) + 1;
      styles.borders(borderId) := border;
      styles.borderMap(border.content) := borderId;
    end if;
    return borderId;
  end;

  procedure setFontContent (
    font  in out nocopy CT_Font
  )
  is
  begin
    string_write(font.content, '<font>');
    string_write(font.content, '<sz val="'||to_char(font.sz)||'"/>');
    string_write(font.content, '<name val="'||to_char(font.name)||'"/>');
    if font.b then
      string_write(font.content, '<b/>');
    end if;
    if font.i then
      string_write(font.content, '<i/>');
    end if;
    if font.color is not null then
      if font.color like 'theme:%' then
        string_write(font.content, '<color theme="'||regexp_substr(font.color,'\d+$')||'"/>');
      else
        string_write(font.content, '<color rgb="'||font.color||'"/>');
      end if;
    end if;
    if font.u is not null then
      string_write(font.content, '<u val="'||font.u||'"/>');
    end if;
    string_write(font.content, '</font>');    
  end;

  function makeFont (
    p_name   in varchar2
  , p_sz     in pls_integer
  , p_b      in boolean default false
  , p_i      in boolean default false
  , p_color  in varchar2 default null
  , p_u      in varchar2 default null
  )
  return CT_Font
  is
    font  CT_Font;
  begin
    font.name := p_name;
    font.sz := p_sz;
    font.b := p_b;
    font.i := p_i;
    
    if p_u is not null then 
      if ExcelTypes.isValidUnderlineStyle(p_u) then
        font.u := p_u;
      else
        error('Invalid underline style: %s', p_u);
      end if;
    end if;
    
    font.color := validateColor(p_color);
    setFontContent(font);
    return font;
  end;
  
  function putFont (
    styles  in out nocopy CT_Stylesheet
  , font    CT_Font
  )
  return pls_integer
  is
    fontId  pls_integer;
  begin
    if styles.fontMap.exists(font.content) then
      fontId := styles.fontMap(font.content);
    else
      fontId := nvl(styles.fonts.last, -1) + 1;
      styles.fonts(fontId) := font;
      styles.fontMap(font.content) := fontId;
    end if;
    return fontId;
  end;
  
  procedure setFillContent (
    fill  in out nocopy CT_Fill
  )
  is
  begin
    string_write(fill.content, '<fill><patternFill patternType="'||fill.patternFill.patternType||'">');
    if fill.patternFill.fgColor is not null then
      string_write(fill.content, '<fgColor rgb="'||fill.patternFill.fgColor||'"/>');
    end if;
    if fill.patternFill.bgColor is not null then
      string_write(fill.content, '<bgColor rgb="'||fill.patternFill.bgColor||'"/>');
    end if;
    string_write(fill.content, '</patternFill></fill>');    
  end;

  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill
  is
    fill  CT_Fill;
  begin
    fill.patternFill.patternType := p_patternType;
    fill.patternFill.fgColor := validateColor(p_fgColor);
    fill.patternFill.bgColor := validateColor(p_bgColor);
    setFillContent(fill);
    return fill;
  end;
  
  function putFill (
    styles  in out nocopy CT_Stylesheet
  , fill    in CT_Fill
  )
  return pls_integer
  is
    fillId  pls_integer;
  begin
    if styles.fillMap.exists(fill.content) then
      fillId := styles.fillMap(fill.content);
    else
      fillId := nvl(styles.fills.last, -1) + 1;
      styles.fills(fillId) := fill;
      styles.fillMap(fill.content) := fillId;
    end if;
    return fillId;
  end;

  procedure setAlignmentContent (
    alignment  in out nocopy CT_CellAlignment
  )
  is
  begin
    if coalesce(alignment.horizontal, alignment.vertical) is not null or alignment.wrapText then
      string_write(alignment.content, '<alignment');
      if alignment.horizontal is not null then
        string_write(alignment.content, ' horizontal="'||alignment.horizontal||'"');
      end if;
      if alignment.vertical is not null then
        string_write(alignment.content, ' vertical="'||alignment.vertical||'"');
      end if;
      if alignment.wrapText then
        string_write(alignment.content, ' wrapText="1"');
      end if;
      string_write(alignment.content, '/>');
    end if;    
  end;

  function makeAlignment (
    p_horizontal  in varchar2 default null
  , p_vertical    in varchar2 default null
  , p_wrapText    in boolean default false
  )
  return CT_CellAlignment
  is
    alignment  CT_CellAlignment;
  begin
    alignment.horizontal := p_horizontal;
    alignment.vertical := p_vertical;
    alignment.wrapText := p_wrapText;
    setAlignmentContent(alignment);
    return alignment;
  end;

  function makeCellXf (
    styles      in out nocopy CT_Stylesheet
  , styleXfId   in pls_integer
  , numFmtCode  in varchar2 default null
  , font        in CT_Font default null
  , fill        in CT_Fill default null
  , border      in CT_Border default null
  , alignment   in CT_CellAlignment default null
  )
  return CT_Xf
  is
    xf  CT_Xf;
  begin
    if styleXfId is not null then
      xf := styles.cellStyleXfs(styleXfId);
      xf.xfId := styleXfId;
      xf.content := null;
    end if;
    
    if numFmtCode is not null then
      xf.numFmtId := putNumfmt(styles, numFmtCode);
    end if;
    if font.content is not null then
      xf.fontId := putFont(styles, font);
    end if;
    if fill.content is not null then
      xf.fillId := putFill(styles, fill);
    end if;
    if border.content is not null then
      xf.borderId := putBorder(styles, border);
    end if;
    
    xf.alignment := alignment;
    
    string_write(xf.content, '<xf');
    string_write(xf.content, ' numFmtId="'||to_char(xf.numFmtId)||'"');
    string_write(xf.content, ' fontId="'||to_char(xf.fontId)||'"');
    string_write(xf.content, ' fillId="'||to_char(xf.fillId)||'"');
    string_write(xf.content, ' borderId="'||to_char(xf.borderId)||'"');
    
    if xf.xfId is not null then
       string_write(xf.content, ' xfId="'||to_char(xf.xfId)||'"');
    end if;
    
    if xf.numFmtId != 0 then
      string_write(xf.content, ' applyNumberFormat="1"');
    end if;
    if xf.fontId != 0 then
      string_write(xf.content, ' applyFont="1"');
    end if;
    if xf.fillId != 0 then
      string_write(xf.content, ' applyFill="1"');
    end if;
    if xf.borderId != 0 then
      string_write(xf.content, ' applyBorder="1"');
    end if;
    
    if xf.alignment.content is not null then
      string_write(xf.content, ' applyAlignment="1">');
      string_write(xf.content, xf.alignment.content);
      string_write(xf.content, '</xf>');
    else
      string_write(xf.content, '/>');
    end if;
    
    return xf;
  end;

  function putCellStyleXf (
    styles  in out nocopy CT_Stylesheet
  , xf      in CT_Xf
  )
  return pls_integer
  is
    xfId  pls_integer;
  begin
    if styles.cellStyleXfMap.exists(xf.content) then
      xfId := styles.cellStyleXfMap(xf.content);
    else
      xfId := nvl(styles.cellStyleXfs.last, -1) + 1;
      styles.cellStyleXfs(xfId) := xf;
      styles.cellStyleXfMap(xf.content) := xfId;
    end if;
    return xfId;
  end;

  function putCellXf (
    styles  in out nocopy CT_Stylesheet
  , xf      in CT_Xf
  )
  return pls_integer
  is
    xfId  pls_integer;
  begin
    if styles.cellXfMap.exists(xf.content) then
      xfId := styles.cellXfMap(xf.content);
    else
      xfId := nvl(styles.cellXfs.last, -1) + 1;
      styles.cellXfs(xfId) := xf;
      styles.cellXfMap(xf.content) := xfId;
    end if;
    return xfId;
  end;

  procedure putNamedCellStyle (
    styles     in out nocopy CT_Stylesheet
  , name       in varchar2
  , xfId       in pls_integer
  , builtinId  in pls_integer
  )
  is
    cellStyle  CT_CellStyle;
  begin
    cellStyle.name := name;
    cellStyle.xfId := xfId;
    cellStyle.builtinId := builtinId;
    styles.cellStyles.extend;
    styles.cellStyles(styles.cellStyles.last) := cellStyle;
  end;  

  function makeCellStyle (
    p_ctxId       in ctxHandle
  , p_numFmtCode  in varchar2 default null
  , p_font        in CT_Font default null
  , p_fill        in CT_Fill default null
  , p_border      in CT_Border default null
  , p_alignment   in CT_CellAlignment default null
  )
  return cellStyleHandle
  is
    xf  CT_Xf := makeCellXf(ctx_cache(p_ctxId).workbook.styles, 0, p_numFmtCode, p_font, p_fill, p_border, p_alignment);
  begin
    return putCellXf(ctx_cache(p_ctxId).workbook.styles, xf);        
  end;

  function getColumnWidth (
    displayWidth in number 
  )
  return binary_double
  is
  begin
    return to_binary_double(
             trunc(
               round(displayWidth * 7 + 5) -- must be an integer number of pixels
               / 7 
               * 256
             ) / 256
           );
  end;

  procedure parseLink (
    link  in out nocopy link_t
  )
  is
    idx      pls_integer := 0;
    tokenId  pls_integer;
    function next_token return pls_integer is
    begin
      idx := idx + 1;
      return to_number(regexp_substr(link.target, '\{(\d+)\}', 1, idx, null, 1));
    end;
    
  begin    
    tokenId := next_token;
    while tokenId is not null loop
      link.tokens(tokenId) := '{'||to_char(tokenId)||'}';
      tokenId := next_token;
    end loop;
  end;

  procedure prepareHyperlink (
    target       in varchar2
  , ctxColumnId  in pls_integer
  , columnList   in out nocopy column_list_t
  --, stylesheet   in CT_Stylesheet
  )
  is
    CTX_TOKEN     constant varchar2(8) := '{' || to_char(ctxColumnId) || '}';
    token         varchar2(8);
    link          link_t;
    location      varchar2(2048);
    columnId      pls_integer;
    --numFmtId      pls_integer;
  begin
    link.target := nvl(target, CTX_TOKEN);
    parseLink(link);
    
    if link.target = CTX_TOKEN then
      link.fmla := 'HYPERLINK(' || enquote(CTX_TOKEN) || ')';
    else
      location := enquote(link.target);
      link.tokens(ctxColumnId) := CTX_TOKEN;
      
      if link.tokens.count != 0 then
        
        columnId := link.tokens.first;
        while columnId is not null loop
          token := link.tokens(columnId);
          --numFmtId := stylesheet.cellXfs(columnList(columnId).xfId).numFmtId;
          if columnId != ctxColumnId and not columnList(columnId).excluded then
            location := replace(location, token, '"&' || token || '&"');
          end if;
          columnId := link.tokens.next(columnId);
        end loop;
      
        -- clean up leading and trailing empty strings
        location := regexp_replace(location, '^""&|&""$');
      
      end if;
      
      link.fmla := 'HYPERLINK(' || location || ',' 
                                || case when columnList(ctxColumnId).supertype = ST_STRING then enquote(CTX_TOKEN) else CTX_TOKEN end || ')';
      
    end if;
    
    columnList(ctxColumnId).link := link;
    columnList(ctxColumnId).hasLink := true;
    
  end;
  
  procedure prepareHyperlinks (
    sheetDef    in out nocopy sheet_definition_t
  --, stylesheet  in CT_Stylesheet
  )
  is
    columnId pls_integer;
  begin
    columnId := sheetDef.columnLinkMap.first;
    while columnId is not null loop
      
      prepareHyperlink(sheetDef.columnLinkMap(columnId), columnId, sheetDef.sqlMetadata.columnList);
    
      columnId := sheetDef.columnLinkMap.next(columnId);
    end loop;
  end;

  function makeHyperlinkFormula (
    link         in link_t
  , ctxRowId     in pls_integer
  , ctxColumnId  in pls_integer
  , columnList   in column_list_t
  , dataMap      in data_map_t
  )
  return varchar2
  is
    fmla  varchar2(8192) := link.fmla;
    columnId  pls_integer;
  begin
    if link.tokens.count != 0 then
      columnId := link.tokens.first;
      while columnId is not null loop
        fmla := replace( fmla
                       , link.tokens(columnId)
                       , case when columnId = ctxColumnId or columnList(columnId).excluded then escapeQuote(dataMap(columnId).varchar2_value) 
                              else columnList(columnId).colRef||to_char(ctxRowId) 
                                end
                       );
        columnId := link.tokens.next(columnId);
      end loop;
    end if;
    dbms_output.put_line(fmla);
    return fmla;
  end;

  function newStylesheet
  return CT_Stylesheet
  is
    styles  CT_Stylesheet;
    dummy   pls_integer;
    xfId    pls_integer;
  begin
    dummy := putFont(styles, makeFont('Calibri', 11));
    dummy := putFill(styles, makePatternFill('none'));
    dummy := putFill(styles, makePatternFill('gray125'));
    dummy := putBorder(styles, makeBorder());
    
    xfId := putCellStyleXf(styles, makeCellXf(styles, null)); -- master cell xf
    dummy := putCellXf(styles, makeCellXf(styles, xfId));
    
    styles.cellStyles := CT_CellStyles();
    putNamedCellStyle(styles, 'Normal', xfId, 0);
    
    return styles;
  end;
  
  function toOADate (dt in date)
  return number 
  deterministic
  is
    output  number := dt - date '1899-12-30';
  begin
    return case when output > 60 then output else output - 1 end;
  end;
  
  function toOADate (ts in timestamp_unconstrained)
  return number 
  deterministic
  is
    dsint   dsinterval_unconstrained := ts - timestamp '1899-12-30 00:00:00';
    output  number;
  begin
    output := extract(day from dsint) + 
              extract(hour from dsint)/24 + 
              extract(minute from dsint)/1440 + 
              extract(second from dsint)/86400;
    return case when output > 60 then output else output - 1 end;
  end;
  
  function timestampRound (
    ts    in timestamp_unconstrained
  , scale in pls_integer default 0
  )
  return timestamp_unconstrained
  deterministic
  is
    seconds  number := extract(second from ts);
  begin
    return ts + numtodsinterval(round(seconds,scale) - seconds, 'second');
  end;
  
  function new_stream
  return stream_t
  is
    stream  stream_t;
  begin
    stream.buf_sz := 0;
    dbms_lob.createtemporary(stream.content, true);
    return stream;
  end;

  function new_stream (
    content  in out nocopy clob
  )
  return stream_t
  is
    stream  stream_t;
  begin
    stream.buf_sz := 0;
    if content is null then
      dbms_lob.createtemporary(content, true);
    end if;
    stream.content := content;
    return stream;
  end;
  
  procedure stream_flush (
    stream  in out nocopy stream_t
  )
  is
  begin
    if stream.buf_sz != 0 then
      dbms_lob.writeappend(stream.content, length(stream.buf), stream.buf);
      stream.buf := null;
      stream.buf_sz := 0;
    end if;
  end;
    
  procedure stream_write (
    stream      in out nocopy stream_t
  , input       in varchar2
  , escape_xml  in boolean default false
  ) 
  is
    chunk     varchar2(32767);
    chunk_sz  pls_integer;
  begin
    if input is not null then
      chunk := case when escape_xml then dbms_xmlgen.convert(input) else input end;
      chunk_sz := lengthb(chunk);
      if stream.buf_sz + chunk_sz <= MAX_BUFFER_SIZE then
        stream.buf := stream.buf || chunk;
        stream.buf_sz := stream.buf_sz + chunk_sz;
      else
        -- flush
        dbms_lob.writeappend(stream.content, length(stream.buf), stream.buf);
        stream.buf := chunk;
        stream.buf_sz := chunk_sz;
      end if;
    end if;
  exception
    when buffer_too_small then
      debug('Switching to CLOB');
      -- flush
      stream_flush(stream);
      -- buffer bypass
      dbms_lob.append(stream.content, dbms_xmlgen.convert(to_clob(input)));
  end;
  
  procedure stream_write_clob (
    stream      in out nocopy stream_t
  , input       in clob
  , max_size    in integer
  , escape_xml  in boolean default false
  )
  is
    buf        varchar2(8191 char);
    amt        integer := 8191;
    pos        integer := 1;
    available  pls_integer := least(dbms_lob.getlength(input), max_size);
  begin
    while available > 0 loop
      amt := least(amt, available);
      dbms_lob.read(input, amt, pos, buf);
      stream_write(stream, case when escape_xml then stripXmlControlChars(buf) else buf end, escape_xml);
      pos := pos + amt;
      available := available - amt;  
    end loop;    
  end;
  
  function put_string (
    ctx in out nocopy context_t
  , str in varchar2
  ) 
  return pls_integer
  is
    idx  pls_integer;
  begin
    ctx.string_cnt := ctx.string_cnt + 1;
    if not ctx.string_map.exists(str) then
      idx := ctx.string_list.count + 1;
      ctx.string_map(str) := idx;
      ctx.string_list.extend;
      ctx.string_list(idx) := str;
    else
      idx := ctx.string_map(str);
    end if;
    return idx;
  end;

  function getCursorNumber (
    p_query in varchar2 
  )
  return integer
  is
    c  integer;
  begin
    c := dbms_sql.open_cursor;
    dbms_sql.parse(c, p_query, dbms_sql.native);
    return c;
  end;

  function getColumnList (
    p_cursor_number  in integer
  , p_excludeSet     in int_set_t
  )
  return column_list_t
  is
    baseColumnList  dbms_sql.desc_tab2;
    columnCount     integer;
    data            data_t;
    columnList      column_list_t := column_list_t();
    COLUMN_DEFAULT  column_t;
    columnItem      column_t;
    columnId        pls_integer := 0;
  begin
    dbms_sql.describe_columns2(p_cursor_number, columnCount, baseColumnList);
    
    for i in 1 .. columnCount loop
      
      columnItem := COLUMN_DEFAULT;
    
      case baseColumnList(i).col_type
      when dbms_sql.VARCHAR2_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.varchar2_value, baseColumnList(i).col_max_len);
        columnItem.supertype := ST_STRING;
      when dbms_sql.CHAR_TYPE then
        dbms_sql.define_column_char(p_cursor_number, i, data.char_value, baseColumnList(i).col_max_len);
        columnItem.supertype := ST_STRING;
      when dbms_sql.NUMBER_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.number_value);
        columnItem.supertype := ST_NUMBER;
      when dbms_sql.DATE_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.date_value);
        columnItem.supertype := ST_DATETIME;
      when dbms_sql.TIMESTAMP_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.ts_value);
        columnItem.supertype := ST_DATETIME;
      when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.tstz_value);
        columnItem.supertype := ST_DATETIME;
      when dbms_sql.CLOB_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.clob_value);
        columnItem.supertype := ST_LOB;
      else
        error('Unsupported data type: %d, for column "%s"', baseColumnList(i).col_type, baseColumnList(i).col_name);
      end case;
      
      columnItem.name := baseColumnList(i).col_name;
      columnItem.type := baseColumnList(i).col_type;
      columnItem.scale := baseColumnList(i).col_scale;
      columnItem.hasLink := false;
      columnItem.excluded := p_excludeSet.exists(i);
      
      if not columnItem.excluded then
        columnId := columnId + 1;
        columnItem.id := columnId;
        columnItem.colRef := base26encode(columnId);
      end if; 
      
      columnList.extend;
      columnList(i) := columnItem;
      
    end loop;
    
    return columnList;
  end;

  procedure prepareCursor (
    meta  in out nocopy sql_metadata_t
  )
  is
    result      integer;
    bind_var    bind_variable_t;
  begin
    
    meta.partitionId := 0;
    meta.r_num := 0;
  
    if meta.cursorNumber is null then
      
      meta.cursorNumber := getCursorNumber(meta.queryString);
      
      -- bind variables
      if meta.bindVariables.count != 0 then
        for i in 1 .. meta.bindVariables.count loop
          bind_var := meta.bindVariables(i);
          case bind_var.value.GetTypeName()
          when 'SYS.VARCHAR2' then 
            dbms_sql.bind_variable(meta.cursorNumber, bind_var.name, bind_var.value.AccessVarchar2());
          when 'SYS.NUMBER' then 
            dbms_sql.bind_variable(meta.cursorNumber, bind_var.name, bind_var.value.AccessNumber());
          when 'SYS.DATE' then 
            dbms_sql.bind_variable(meta.cursorNumber, bind_var.name, bind_var.value.AccessDate());
          end case;        
        end loop;
      end if;
      
      debug('execute cursor');
      result := dbms_sql.execute(meta.cursorNumber);
      
    end if;
    
    meta.columnList := getColumnList(meta.cursorNumber, meta.excludeSet);
    
    for i in 1 .. meta.columnList.count loop
      meta.columnMap(meta.columnList(i).name) := i;
      if not meta.columnList(i).excluded then
        meta.visibleColumnSet(i) := meta.columnList(i).colRef;
      end if;
    end loop;
    
  end;

  function getRelativePath (
    pathName1  in varchar2
  , pathName2  in varchar2 
  )
  return varchar2
  is
    type path_t is table of varchar2(256);
    
    function tokenize (pathName in varchar2) return path_t;
    
    path1   path_t := tokenize(pathName1);
    path2   path_t := tokenize(pathName2);
    idx     pls_integer := 1;
    cnt     pls_integer := 0;
    output  varchar2(256);

    function tokenize (
      pathName in varchar2
    )
    return path_t
    is
      path  path_t := path_t();
      step  varchar2(256);
      p1    pls_integer := 1;
      p2    pls_integer;  
    begin
      if pathName is not null then
        loop
          p2 := instr(pathName, '/', p1);
          if p2 = 0 then
            step := substr(pathName, p1);
          else
            step := substr(pathName, p1, p2-p1);    
            p1 := p2 + 1;
          end if;
          path.extend;
          path(path.last) := step;
          exit when p2 = 0;
        end loop;
      end if;
      return path; 
    end;
    
  begin
    
    while idx < path1.count loop
      if path1(idx) != path2(idx) then
        cnt := path1.count - idx;
        exit;
      end if;
      idx := idx + 1;
    end loop;
    
    for i in 1 .. cnt loop
      output := output || '../';
    end loop;
    
    output := output || path2(idx);
    for i in idx + 1 .. path2.count loop
      output := output || '/' || path2(i);
    end loop;
    
    return output;

  end;  

  procedure addPart (
    ctx   in out nocopy context_t
  , part  in part_t
  )
  is
    idx  pls_integer;
  begin
    ctx.pck.parts.extend;
    idx := ctx.pck.parts.last;
    ctx.pck.parts(idx) := part;
    ctx.pck.partIndices(part.name) := idx;
  end;

  procedure addPart (
    ctx          in out nocopy context_t
  , name         in varchar2
  , contentType  in varchar2
  , content      in clob
  )
  is
    idx  pls_integer;
  begin
    ctx.pck.parts.extend;
    idx := ctx.pck.parts.last;
    ctx.pck.parts(idx).name := name;
    ctx.pck.parts(idx).contentType := contentType;
    ctx.pck.parts(idx).content := content;
    ctx.pck.parts(idx).rels := CT_Relationships();
    ctx.pck.partIndices(name) := idx;
  end;

  procedure addPart (
    ctx          in out nocopy context_t
  , name         in varchar2
  , contentType  in varchar2
  , contentBin   in blob
  )
  is
    idx  pls_integer;
  begin
    ctx.pck.parts.extend;
    idx := ctx.pck.parts.last;
    ctx.pck.parts(idx).name := name;
    ctx.pck.parts(idx).contentType := contentType;
    ctx.pck.parts(idx).contentBin := contentBin;
    ctx.pck.parts(idx).isBinary := true;
    ctx.pck.parts(idx).rels := CT_Relationships();
    ctx.pck.partIndices(name) := idx;
  end;

  function addRelationship (
    part    in out nocopy part_t
  , type    in varchar2
  , target  in varchar2
  )
  return varchar2
  is
    i          pls_integer;
    rId        varchar2(256);
    relTarget  varchar2(256) := getRelativePath(part.name, target);
  begin
    part.rels.extend;
    i := part.rels.last;
    rId := 'rId'||to_char(i);
    part.rels(i).id := rId;
    part.rels(i).type := type;
    part.rels(i).target := relTarget;
    return rId;
  end;
  
  procedure addRelationship (
    ctx       in out nocopy context_t
  , partName  in varchar2
  , type      in varchar2
  , target    in varchar2
  )
  is
    i          pls_integer;
    j          pls_integer;
    relTarget  varchar2(256) := getRelativePath(partName, target);
  begin
    if partName is not null then
      i := ctx.pck.partIndices(partName);
      ctx.pck.parts(i).rels.extend;
      j := ctx.pck.parts(i).rels.last;
      ctx.pck.parts(i).rels(j).id := 'rId'||to_char(j);
      ctx.pck.parts(i).rels(j).type := type;
      ctx.pck.parts(i).rels(j).target := relTarget;
    else
      ctx.pck.rels.extend;
      i := ctx.pck.rels.last;
      ctx.pck.rels(i).id := 'rId'||to_char(i);
      ctx.pck.rels(i).type := type;
      ctx.pck.rels(i).target := relTarget;      
    end if;
  end;

  /*
  procedure addSheet (
    ctx      in out nocopy context_t
  , name     in varchar2
  , content  in clob
  , autoFilterRange  in range_t default null
  )
  is
    sheet     CT_Sheet;
    i         pls_integer;
  begin
    ctx.workbook.sheets.extend;
    i := ctx.workbook.sheets.last;
    sheet.partName := 'xl/worksheets/sheet'||to_char(i)||'.xml';
    sheet.name := name;
    sheet.sheetId := i;
    if autoFilterRange.expr is not null then
      sheet.filterRange := autoFilterRange;
      ctx.workbook.hasDefinedNames := true;
    end if;
    ctx.workbook.sheets(i) := sheet;
    addPart(ctx, sheet.partName, MT_WORKSHEET, content);
  end;
  */

  function addTable (
    ctx              in out nocopy context_t
  , tableRange       in range_t
  , showHeader       in boolean
  , tableAutoFilter  in boolean
  , tableStyleName   in varchar2
  --, columnList       in column_list_t
  , columnMap        in sheet_column_map_t
  , tableName        in varchar2 default null
  , isEmpty          in boolean default false
  )
  return pls_integer
  is
    tab  CT_Table;
  begin
    tab.id := nvl(ctx.workbook.tables.last, 0) + 1;
    tab.name := nvl(tableName, 'Table'||to_char(tab.id));
    -- if the table is declared over an empty dataset, extends its range by one row down to make it legal in Excel
    if isEmpty then
      tab.ref := makeRange(tableRange.start_ref.c, tableRange.start_ref.r, tableRange.end_ref.c, tableRange.end_ref.r + 1);
    else
      tab.ref := tableRange;
    end if;
    
    tab.showHeader := nvl(showHeader, false);
    tab.autoFilter := nvl(tableAutoFilter, false);
    tab.styleName := tableStyleName;
    tab.partName := 'xl/tables/table'||to_char(tab.id)
                                     ||case when ctx.fileType = FILE_XLSB then '.bin' else '.xml' end;
    tab.cols := CT_TableColumns();
    /*
    tab.cols.extend(columnList.count);
    for i in 1 .. columnList.count loop
      tab.cols(i).id := i;
      tab.cols(i).name := columnList(i).name;
    end loop;
    */
    tab.cols.extend(columnMap.count);
    for i in 1 .. columnMap.count loop
      tab.cols(i).id := i;
      tab.cols(i).name := columnMap(i).header;
    end loop;
    
    ctx.workbook.tables(tab.id) := tab;
    return tab.id;
  end;
  
  procedure createRelPart (
    ctx       in out nocopy context_t
  , partName  in varchar2
  , rels      in CT_Relationships
  )
  is
    stream       stream_t;
    relPartName  varchar2(256);
  begin
    if rels.count != 0 then
      stream := new_stream();
      stream_write(stream, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      for i in 1 .. rels.count loop
        stream_write(stream, '<Relationship Id="'||rels(i).Id||'" Type="'||rels(i).Type||'" Target="'||rels(i).Target||'"/>');
      end loop;
      stream_write(stream, '</Relationships>');
      stream_flush(stream);
      relPartName := substr(partName, 1, instr(partName, '/', -1)) || 
                     '_rels/' || 
                     substr(partName, instr(partName, '/', -1) + 1) || 
                     '.rels' ;
      addPart(ctx, relPartName, null, stream.content);
      debug(xmltype(stream.content).getstringval(1,2));
    end if;
  end;
  
  procedure createRels (
    ctx   in out nocopy context_t
  )
  is
  begin
    -- generate package-level relationships
    createRelPart(ctx, null, ctx.pck.rels);
    -- generate part relationships
    for i in 1 .. ctx.pck.parts.count loop
      createRelPart(ctx, ctx.pck.parts(i).name, ctx.pck.parts(i).rels);
    end loop;
  end;
  
  function new_workbook
  return CT_Workbook
  is
    wb  CT_Workbook;
  begin
    wb.sheets := CT_Sheets();
    wb.styles := newStylesheet();
    return wb;
  end;
  
  procedure addDefaultStyles (
    styles  in out nocopy CT_Stylesheet
  )
  is
    styleXfId   pls_integer;
    xfId         pls_integer;
    defaultFont  CT_Font := styles.fonts(0);
    hlinkFont    CT_Font;
  begin
    if styles.hasHlink then
      -- new hyperlink font derived from default
      hlinkFont := makeFont(defaultFont.name, defaultFont.sz, defaultFont.b, defaultFont.i, 'theme:10', 'single');
      -- new master cell xf using this font
      styleXfId := putCellStyleXf(styles, makeCellXf(styles, null, font => hlinkFont)); -- master cell xf
      styles.hlinkXfId := styleXfId;
      -- new cell xf derived from master (moved to createWorksheet)
      -- styles.hlinkXfId := putCellXf(styles, makeCellXf(styles, styleXfId));
      -- new named cell style for builtinId 8 (= hyperlink style)
      putNamedCellStyle(styles, 'Hyperlink', styleXfId, 8);
    end if;
  end;

  procedure createStylesheet (
    ctx       in out nocopy context_t
  , styles    in CT_Stylesheet
  , partName  in varchar2
  )
  is
    stream  stream_t;
    
  begin
    stream := new_stream();
    stream_write(stream, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    
    -- numFmts
    if styles.numFmts.count != 0 then
      stream_write(stream, '<numFmts count="'||to_char(styles.numFmts.count)||'">');
      for numFmtId in styles.numFmts.first .. styles.numFmts.last loop
        stream_write(stream, '<numFmt numFmtId="'||to_char(numFmtId)||'" formatCode="'||dbms_xmlgen.convert(styles.numFmts(numFmtId))||'"/>');
      end loop;
      stream_write(stream, '</numFmts>');
    end if;
    
    -- fonts
    if styles.fonts.count != 0 then
      stream_write(stream, '<fonts count="'||to_char(styles.fonts.count)||'">');
      for fontId in styles.fonts.first .. styles.fonts.last loop
        stream_write(stream, styles.fonts(fontId).content);
      end loop;
      stream_write(stream, '</fonts>');
    end if;
    
    -- fills
    if styles.fills.count != 0 then
      stream_write(stream, '<fills count="'||to_char(styles.fills.count)||'">');
      for fillId in styles.fills.first .. styles.fills.last loop
        stream_write(stream, styles.fills(fillId).content);
      end loop;
      stream_write(stream, '</fills>');
    end if;
    
    -- borders
    if styles.borders.count != 0 then
      stream_write(stream, '<borders count="'||to_char(styles.borders.count)||'">');
      for borderId in styles.borders.first .. styles.borders.last loop
        stream_write(stream, styles.borders(borderId).content);
      end loop;
      stream_write(stream, '</borders>');
    end if;
    
    -- cellStyleXfs
    if styles.cellStyleXfs.count != 0 then
      stream_write(stream, '<cellStyleXfs count="'||to_char(styles.cellStyleXfs.count)||'">');
      for i in styles.cellStyleXfs.first .. styles.cellStyleXfs.last loop
        stream_write(stream, styles.cellStyleXfs(i).content);
      end loop;
      stream_write(stream, '</cellStyleXfs>');
    end if;

    -- cellXfs
    if styles.cellXfs.count != 0 then
      stream_write(stream, '<cellXfs count="'||to_char(styles.cellXfs.count)||'">');
      for xfId in styles.cellXfs.first .. styles.cellXfs.last loop
        stream_write(stream, styles.cellXfs(xfId).content);
      end loop;
      stream_write(stream, '</cellXfs>');
    end if;
    
    -- cellStyles
    if styles.cellStyles.count != 0 then
      stream_write(stream, '<cellStyles count="'||to_char(styles.cellStyles.count)||'">');
      for i in 1 .. styles.cellStyles.count loop
        stream_write(stream, '<cellStyle name="' || dbms_xmlgen.convert(styles.cellStyles(i).name) || 
                                      '" xfId="' || to_char(styles.cellStyles(i).xfId) || 
                                      '" builtinId="' || to_char(styles.cellStyles(i).builtinId) || 
                                      '"/>');
      end loop;
      stream_write(stream, '</cellStyles>');
    end if;
    
    -- dxfs
    stream_write(stream, '<dxfs count="0"/>');
    
    stream_write(stream, '</styleSheet>');
    stream_flush(stream);
    debug(xmltype(stream.content).getstringval(1,2));
    addPart(ctx, partName, MT_STYLES, stream.content);
    
  end;

  procedure createStylesheetBin (
    ctx       in out nocopy context_t
  , styles    in CT_Stylesheet
  , partName  in varchar2
  )
  is
    stream  xutl_xlsb.Stream_T := xutl_xlsb.new_stream();
  begin
    
    xutl_xlsb.put_simple_record(stream, 278); -- BrtBeginStyleSheet
    
    -- numFmts
    if styles.numFmts.count != 0 then
      xutl_xlsb.put_simple_record(stream, 615, int2raw(styles.numFmts.count)); -- BrtBeginFmts
      for numFmtId in styles.numFmts.first .. styles.numFmts.last loop
        -- BrtFmt
        xutl_xlsb.put_NumFmt(stream, numFmtId, styles.numFmts(numFmtId));
      end loop;
      xutl_xlsb.put_simple_record(stream, 616); -- BrtEndFmts
    end if;
    
    -- fonts
    if styles.fonts.count != 0 then
      xutl_xlsb.put_simple_record(stream, 611, int2raw(styles.fonts.count)); -- BrtBeginFonts
      for fontId in styles.fonts.first .. styles.fonts.last loop
        -- BrtFont
        xutl_xlsb.put_Font(stream, styles.fonts(fontId));
      end loop;
      xutl_xlsb.put_simple_record(stream, 612); -- BrtEndFonts
    end if;
    
    -- fills
    if styles.fills.count != 0 then
      xutl_xlsb.put_simple_record(stream, 603, int2raw(styles.fills.count)); -- BrtBeginFills
      for fillId in styles.fills.first .. styles.fills.last loop
        -- BrtFill
        xutl_xlsb.put_PatternFill(stream, styles.fills(fillId).patternFill);
      end loop;
      xutl_xlsb.put_simple_record(stream, 604); -- BrtEndFills
    end if;
    
    -- borders
    if styles.borders.count != 0 then
      xutl_xlsb.put_simple_record(stream, 613, int2raw(styles.borders.count)); -- BrtBeginBorders
      for borderId in styles.borders.first .. styles.borders.last loop
        -- BrtBorder
        xutl_xlsb.put_Border(stream, styles.borders(borderId));
      end loop;
      xutl_xlsb.put_simple_record(stream, 614); -- BrtEndBorders
    end if;
    
    -- cellStyleXfs
    xutl_xlsb.put_simple_record(stream, 626, int2raw(styles.cellStyleXfs.count));  -- BrtBeginCellStyleXFs
    for xfId in styles.cellXfs.first .. styles.cellXfs.last loop
      -- BrtXF
      xutl_xlsb.put_XF(stream 
                     --, xfId       => styles.cellXfs(xfId).xfId
                     , numFmtId   => styles.cellXfs(xfId).numFmtId
                     , fontId     => styles.cellXfs(xfId).fontId
                     , fillId     => styles.cellXfs(xfId).fillId
                     , borderId   => styles.cellXfs(xfId).borderId
                     , hAlignment => styles.cellXfs(xfId).alignment.horizontal
                     , vAlignment => styles.cellXfs(xfId).alignment.vertical
                     , wrapText   => styles.cellXfs(xfId).alignment.wrapText
                     );
    end loop;
    xutl_xlsb.put_simple_record(stream, 627);  -- BrtEndCellStyleXFs
    
    -- cellXfs
    if styles.cellXfs.count != 0 then
      xutl_xlsb.put_simple_record(stream, 617, int2raw(styles.cellXfs.count));  -- BrtBeginCellXFs
      for xfId in styles.cellXfs.first .. styles.cellXfs.last loop
        -- BrtXF
        xutl_xlsb.put_XF(stream 
                       , xfId       => styles.cellXfs(xfId).xfId
                       , numFmtId   => styles.cellXfs(xfId).numFmtId
                       , fontId     => styles.cellXfs(xfId).fontId
                       , fillId     => styles.cellXfs(xfId).fillId
                       , borderId   => styles.cellXfs(xfId).borderId
                       , hAlignment => styles.cellXfs(xfId).alignment.horizontal
                       , vAlignment => styles.cellXfs(xfId).alignment.vertical
                       , wrapText   => styles.cellXfs(xfId).alignment.wrapText
                       );
      end loop;
      xutl_xlsb.put_simple_record(stream, 618);  -- BrtEndCellXFs
    end if;
    
    -- cellStyles
    xutl_xlsb.put_simple_record(stream, 619, int2raw(styles.cellStyles.count));  -- BrtBeginStyles
    for i in 1 .. styles.cellStyles.count loop
      xutl_xlsb.put_BuiltInStyle(stream, styles.cellStyles(i).builtinId, styles.cellStyles(i).name, styles.cellStyles(i).xfId);  -- BrtStyle
    end loop;
    xutl_xlsb.put_simple_record(stream, 620);  -- BrtEndStyles    
    
    -- dxfs
    xutl_xlsb.put_simple_record(stream, 505, int2raw(0));  -- BrtBeginDXFs
    xutl_xlsb.put_simple_record(stream, 506);  -- BrtEndDXFs 
    
    -- tableStyles?
    
    xutl_xlsb.put_simple_record(stream, 279); -- BrtEndStyleSheet
    xutl_xlsb.flush_stream(stream);
    addPart(ctx, partName, MT_STYLES_BIN, stream.content);
    
  end;

  procedure createSharedStrings (
    ctx   in out nocopy context_t
  )
  is
    stream  stream_t;
  begin
    debug('start create sst');
    if ctx.string_cnt != 0 then
      stream := new_stream();
      stream_write(stream, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'||to_char(ctx.string_cnt)||'" uniqueCount="'||to_char(ctx.string_map.count)||'">');
      for i in 1 .. ctx.string_list.count loop
        stream_write(stream, '<si><t>');
        stream_write(stream, ctx.string_list(i), escape_xml => true);
        stream_write(stream, '</t></si>');
      end loop;
      stream_write(stream, '</sst>');
      stream_flush(stream);
      addPart(ctx, 'xl/sharedStrings.xml', MT_SHAREDSTRINGS, stream.content);
    end if;
    debug('end create sst');
  end;

  procedure createSharedStringsBin (
    ctx   in out nocopy context_t
  )
  is
    stream  xutl_xlsb.Stream_T;
  begin
    if ctx.string_cnt != 0 then
      stream := xutl_xlsb.new_stream();
      xutl_xlsb.put_BeginSst(stream, ctx.string_cnt, ctx.string_map.count); -- BrtBeginSst
      for i in 1 .. ctx.string_list.count loop
        xutl_xlsb.put_SSTItem(stream, ctx.string_list(i));
      end loop;
      xutl_xlsb.put_simple_record(stream, 160); -- BrtEndSst
      xutl_xlsb.flush_stream(stream);
      addPart(ctx, 'xl/sharedStrings.bin', MT_SHAREDSTRINGS_BIN, stream.content);
    end if;
  end;
  
  procedure createWorksheetImpl (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t
  )
  is
    data            data_t;
    r               data_row_t;               
    nrows           integer;
    rowIdx          integer := 0;
    isEmpty         boolean := true;
    cellRef         varchar2(10);
    sst_idx         pls_integer;
    stream          stream_t;
    
    columnId        pls_integer;
    columnHeader    varchar2(256);
    
    cellXfId        pls_integer;
    cellHasLink     boolean;

    sheetRange      range_t;
    tableId         pls_integer;
    rId             varchar2(256);
    
    part            part_t;
    sheet           CT_Sheet;
    
    partitionStart  pls_integer;
    partitionStop   pls_integer;
    
    link            link_t;
    
    function makeHyperlinkCellContent (columnId in pls_integer, strValue in varchar2) return varchar2 is
    begin
      return '<c r="'||cellRef||'" '
                     ||case when sd.sqlMetadata.columnList(columnId).supertype = ST_STRING then 't="str" ' end
                     ||'s="'||to_char(cellXfId)||'"><f>'
                     ||dbms_xmlgen.convert(makeHyperlinkFormula(link, rowIdx, columnId, sd.sqlMetadata.columnList, r.dataMap))
                     ||'</f><v>'||dbms_xmlgen.convert(strValue)||'</v></c>' ;
    end;
    
  begin

    -- prefetch
    nrows := dbms_sql.fetch_rows(sd.sqlMetadata.cursorNumber);
    
    if nrows != 0 or ( nrows = 0 and sd.sqlMetadata.partitionId = 0 ) then
      
      isEmpty := (nrows = 0);
      stream := new_stream();  
      stream_write(stream, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
      
      if sd.tabColor is not null then
        stream_write(stream, '<sheetPr><tabColor rgb="'||sd.tabColor||'"/></sheetPr>');
      end if;
      
      if sd.header.show and sd.header.isFrozen then
        stream_write(stream, '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>');
      end if;
      
      -- columns
      if sd.columnMap.count != 0 then
        stream_write(stream, '<cols>');
        for i in 1 .. sd.columnMap.count loop
          if sd.columnMap(i).width is not null then
            stream_write(stream, '<col min="'||to_char(i)||'" max="'||to_char(i)||'" width="'||to_char(getColumnWidth(sd.columnMap(i).width),'TM9')||'" customWidth="1"/>');
          end if;
        end loop;
        stream_write(stream, '</cols>');
      end if;
      
      stream_write(stream, '<sheetData>');
      
      -- header row
      if sd.header.show then
        rowIdx := rowIdx + 1;
        stream_write(stream, '<row r="'||to_char(rowIdx)||'">');
        for i in 1 .. sd.sqlMetadata.columnList.count loop
          if not sd.sqlMetadata.columnList(i).excluded then
            columnId := sd.sqlMetadata.columnList(i).id;
            columnHeader := sd.columnMap(columnId).header;
            sst_idx := put_string(ctx, columnHeader);
            
            cellRef := sd.sqlMetadata.columnList(i).colRef||to_char(rowIdx);
            stream_write(stream, '<c r="'||cellRef||'" t="s"'||
                                 case when sd.header.xfId is not null then ' s="'||to_char(sd.header.xfId)||'"' end || 
                                 '><v>'||to_char(sst_idx - 1)||'</v></c>');
          end if;
        end loop;
        stream_write(stream, '</row>');
      end if;
      
      partitionStart := sd.sqlMetadata.r_num + nrows;
      partitionStop := partitionStart + sd.sqlMetadata.partitionSize - 1;
      
      -- data rows
      while nrows != 0 loop
        
        rowIdx := rowIdx + 1;
        stream_write(stream, '<row r="'||to_char(rowIdx)||'">');
        
        r.dataMap.delete;
        -- read current row
        for i in 1 .. sd.sqlMetadata.columnList.count loop
                    
          data := null;

          case sd.sqlMetadata.columnList(i).type
          when dbms_sql.VARCHAR2_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.varchar2_value);
            data.varchar2_value := stripXmlControlChars(data.varchar2_value);
            
          when dbms_sql.CHAR_TYPE then
            dbms_sql.column_value_char(sd.sqlMetadata.cursorNumber, i, data.char_value);
            data.varchar2_value := stripXmlControlChars(rtrim(data.char_value));
            
          when dbms_sql.NUMBER_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.number_value);
            if sd.sqlMetadata.columnList(i).scale between -84 and 0 then
              data.varchar2_value := to_char(data.number_value);
            else
              data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
            end if;
            
          when dbms_sql.DATE_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.date_value);
            data.number_value := toOADate(dt => data.date_value);
            data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
            
          when dbms_sql.TIMESTAMP_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.ts_value);
            data.ts_value := timestampRound(data.ts_value, 3);
            data.number_value := toOADate(ts => data.ts_value);
            data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
            
          when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.tstz_value);
            data.tstz_value := timestampRound(data.tstz_value, 3);
            data.number_value := toOADate(ts => data.tstz_value);
            data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
            
          when dbms_sql.CLOB_TYPE then      
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.clob_value);
            
          end case;
          
          r.dataMap(i) := data;
          
        end loop;

        for i in 1 .. sd.sqlMetadata.columnList.count loop
          
          if not sd.sqlMetadata.columnList(i).excluded then
          
            data := r.dataMap(i);
            
            cellRef := sd.sqlMetadata.columnList(i).colRef||to_char(rowIdx);
            cellXfId := sd.sqlMetadata.columnList(i).xfId;
            
            cellHasLink := sd.sqlMetadata.columnList(i).hasLink;
            if cellHasLink then
              link := sd.sqlMetadata.columnList(i).link;
            end if;
            
            case sd.sqlMetadata.columnList(i).supertype
            when ST_STRING then
              if data.varchar2_value is not null then
                if not cellHasLink then
                  sst_idx := put_string(ctx, data.varchar2_value);
                  stream_write(stream, '<c r="'||cellRef||'" t="s"><v>'||to_char(sst_idx - 1)||'</v></c>');
                else
                  stream_write(stream, makeHyperlinkCellContent(i, escapeQuote(data.varchar2_value)));
                end if;
              end if;
              
            when ST_NUMBER then
              if not cellHasLink then
                stream_write(stream, '<c r="'||cellRef
                    ||case when cellXfId != 0 then '" s="'||to_char(cellXfId) end
                    ||'"><v>'||data.varchar2_value||'</v></c>');
              else
                stream_write(stream, makeHyperlinkCellContent(i, data.varchar2_value));
              end if;
              
            when ST_DATETIME then
              if not cellHasLink then
                stream_write(stream, '<c r="'||cellRef||'" s="'||to_char(cellXfId)||'"><v>'||data.varchar2_value||'</v></c>');
              else
                stream_write(stream, makeHyperlinkCellContent(i, data.varchar2_value));
              end if;
              
            when ST_LOB then      
              if data.clob_value is not null and dbms_lob.getlength(data.clob_value) != 0 then
                -- try conversion to VARCHAR2
                begin
                  data.varchar2_value := to_char(data.clob_value);
                  sst_idx := put_string(ctx, stripXmlControlChars(data.varchar2_value));
                  stream_write(stream, '<c r="'||cellRef||'" t="s"><v>'||to_char(sst_idx - 1)||'</v></c>');
                exception
                  when value_error then
                    -- stream CLOB content as inlineStr, up to 32767 chars
                    stream_write(stream, '<c r="'||cellRef||'" t="inlineStr"><is><t>');
                    stream_write_clob(stream, data.clob_value, 32767, true);
                    stream_write(stream, '</t></is></c>');
                end;
              end if;
              
            end case;
          
          end if;
          
        end loop;
        
        stream_write(stream, '</row>');
        
        sd.sqlMetadata.r_num := sd.sqlMetadata.r_num + 1;
        
        if rowIdx = MAX_ROW_NUMBER then
          if not sd.sqlMetadata.partitionBySize then
            -- force closing cursor
            nrows := 0;
          end if;
          exit;
        end if;
        
        exit when sd.sqlMetadata.r_num = partitionStop;
        
        -- fetch next row
        nrows := dbms_sql.fetch_rows(sd.sqlMetadata.cursorNumber);
          
      end loop;
      
      debug('end fetch');
               
      stream_write(stream, '</sheetData>');
      
      sheetRange := makeRange(sd.sqlMetadata.columnList(sd.sqlMetadata.visibleColumnSet.first).colRef
                            , 1
                            , sd.sqlMetadata.columnList(sd.sqlMetadata.visibleColumnSet.last).colRef
                            , rowIdx);
      
      -- autoFilter
      if sd.header.show and sd.header.autoFilter then
        if not sd.formatAsTable then
          sheet.filterRange := sheetRange;
          ctx.workbook.hasDefinedNames := true;
          stream_write(stream, '<autoFilter ref="'||getRangeExpr(sheetRange)||'"/>');
        end if;
      end if;
         
      -- new sheet
      sd.sqlMetadata.partitionId := sd.sqlMetadata.partitionId + 1;
      ctx.workbook.sheets.extend;
      sheet.sheetId := ctx.workbook.sheets.last;
      sheet.name := sd.sheetName;
      if sd.sqlMetadata.partitionBySize then
        sheet.name := replace(sheet.name, '${PNUM}', to_char(sd.sqlMetadata.partitionId));
        sheet.name := replace(sheet.name, '${PSTART}', to_char(partitionStart));
        sheet.name := replace(sheet.name, '${PSTOP}', to_char(sd.sqlMetadata.r_num));
      end if;
      
      -- check name validity
      if translate(sheet.name, '_\/*?:[]', '_') != sheet.name 
         or substr(sheet.name, 1, 1) = '''' 
         or substr(sheet.name, -1) = ''''
         or length(sheet.name) > 31 
      then
        error('Invalid sheet name: %s', sheet.name);
      end if;
      
      -- check name uniqueness
      if ctx.workbook.sheetMap.exists(sheet.name) then
        error('Duplicate sheet name: %s', sheet.name);
      end if;
      
      sheet.partName := 'xl/worksheets/sheet'||to_char(sheet.sheetId)||'.xml';
      sheet.tableParts := CT_TableParts();

      -- new sheet part
      part.name := sheet.partName;
      part.contentType := MT_WORKSHEET;
      part.rels := CT_Relationships();
      
      if sd.formatAsTable then
        tableId := addTable(ctx, sheetRange, sd.header.show, sd.header.autoFilter, sd.tableStyle, sd.columnMap, null, isEmpty);
        sheet.tableParts.extend;
        sheet.tableParts(sheet.tableParts.last) := tableId;
        
        -- table parts
        if sheet.tableParts.count != 0 then
          stream_write(stream, '<tableParts count="'||to_char(sheet.tableParts.count)||'">');
          for i in 1 .. sheet.tableParts.count loop
            rId := addRelationship(part, RS_TABLE, ctx.workbook.tables(sheet.tableParts(i)).partName);
            stream_write(stream, '<tablePart r:id="'||rId||'"/>');
          end loop;
          stream_write(stream, '</tableParts>');
        end if;
      end if;
      
      stream_write(stream, '</worksheet>');
      stream_flush(stream);
      
      part.content := stream.content;
      
      -- add sheet to workbook
      ctx.workbook.sheets(sheet.sheetId) := sheet;
      ctx.workbook.sheetMap(sheet.name) := sheet.sheetId;
      
      -- add sheet part to package
      addPart(ctx, part);
    
    end if;
      
    if nrows = 0 then
      debug('close cursor');
      dbms_sql.close_cursor(sd.sqlMetadata.cursorNumber);
    end if;

  end;

  procedure createWorksheetBinImpl (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t
  )
  is
    data            data_t;
    nrows           integer;
    rowIdx          integer := 0;
    isEmpty         boolean := true;
    sst_idx         pls_integer;
    stream          xutl_xlsb.Stream_T;
    
    columnId        pls_integer;
    columnHeader    varchar2(256);
    
    cellXfId        pls_integer;

    sheetRange      range_t;
    tableId         pls_integer;
    rId             varchar2(256);
    
    part            part_t;
    sheet           CT_Sheet;
    
    partitionStart  pls_integer;
    partitionStop   pls_integer;
    
  begin
    
    -- prefetch
    nrows := dbms_sql.fetch_rows(sd.sqlMetadata.cursorNumber);
    
    if nrows != 0 or ( nrows = 0 and sd.sqlMetadata.partitionId = 0 ) then
    
      isEmpty := (nrows = 0);
      stream := xutl_xlsb.new_stream();  
      
      xutl_xlsb.put_simple_record(stream, 129);  -- BrtBeginSheet
      
      if sd.tabColor is not null then
        xutl_xlsb.put_WsProp(stream, sd.tabColor);
      end if;
      

      if sd.header.show and sd.header.isFrozen then
        xutl_xlsb.put_simple_record(stream, 133);  -- BrtBeginWsViews
        xutl_xlsb.put_BeginWsView(stream);  -- BrtBeginWsView
        -- BrtPane : 
        xutl_xlsb.put_FrozenPane(stream
                               , numRows => 1  -- num of frozen rows (ySplit)
                               , numCols => 0  -- num of frozen columns  (xSplit)
                               , topRow  => 1  -- first row of bottom-right pane
                               , leftCol => 0  -- first column of bottom-right pane
                               );
        xutl_xlsb.put_simple_record(stream, 138);  -- BrtEndWsView
        xutl_xlsb.put_simple_record(stream, 134);  -- BrtEndWsViews
      end if;

      -- columns
      if sd.columnMap.count != 0 then
        xutl_xlsb.put_simple_record(stream, 390);  -- BrtBeginColInfos
        for i in 1 .. sd.columnMap.count loop
          if sd.columnMap(i).width is not null then
            xutl_xlsb.put_ColInfo(stream, i - 1, getColumnWidth(sd.columnMap(i).width) * 256);
          end if;
        end loop;
        xutl_xlsb.put_simple_record(stream, 391);  -- BrtEndColInfos
      end if;
      
      xutl_xlsb.put_simple_record(stream, 145);  -- BrtBeginSheetData
      
      -- header row
      if sd.header.show then
        rowIdx := rowIdx + 1;
        xutl_xlsb.put_RowHdr(stream, rowIdx - 1);
        for i in 1 .. sd.sqlMetadata.columnList.count loop
          if not sd.sqlMetadata.columnList(i).excluded then
            columnId := sd.sqlMetadata.columnList(i).id;
            columnHeader := sd.columnMap(columnId).header;
            sst_idx := put_string(ctx, columnHeader);
            xutl_xlsb.put_CellIsst(stream
                                 , colIndex => i - 1
                                 , styleRef => nvl(sd.header.xfId, 0)
                                 , isst     => sst_idx - 1
                                 );
          end if;
        end loop;
      end if;
      
      partitionStart := sd.sqlMetadata.r_num + nrows;
      partitionStop := partitionStart + sd.sqlMetadata.partitionSize - 1;
      
      -- data rows
      while nrows != 0 loop
        
        rowIdx := rowIdx + 1;
        xutl_xlsb.put_RowHdr(stream, rowIdx - 1);
        
        for i in 1 .. sd.sqlMetadata.columnList.count loop
          
          cellXfId := sd.sqlMetadata.columnList(i).xfId;
          
          case sd.sqlMetadata.columnList(i).type
          when dbms_sql.VARCHAR2_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.varchar2_value);
            if data.varchar2_value is not null then
              sst_idx := put_string(ctx, data.varchar2_value);
              xutl_xlsb.put_CellIsst(stream, i-1, 0, sst_idx-1);
            end if;
            
          when dbms_sql.CHAR_TYPE then
            dbms_sql.column_value_char(sd.sqlMetadata.cursorNumber, i, data.char_value);
            if data.char_value is not null then
              data.varchar2_value := rtrim(data.char_value);
              sst_idx := put_string(ctx, data.varchar2_value);
              xutl_xlsb.put_CellIsst(stream, i-1, 0, sst_idx-1);
            end if;
            
          when dbms_sql.NUMBER_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.number_value);
            xutl_xlsb.put_CellNumber(stream, i-1, cellXfId, data.number_value);
            
          when dbms_sql.DATE_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.date_value);
            xutl_xlsb.put_CellNumber(stream, i-1, cellXfId, toOADate(dt => data.date_value));
            
          when dbms_sql.TIMESTAMP_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.ts_value);
            data.ts_value := timestampRound(data.ts_value, 3);
            xutl_xlsb.put_CellNumber(stream, i-1, cellXfId, toOADate(ts => data.ts_value));
            
          when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.tstz_value);
            data.ts_value := timestampRound(data.tstz_value, 3);
            xutl_xlsb.put_CellNumber(stream, i-1, cellXfId, toOADate(ts => data.tstz_value));
            
          when dbms_sql.CLOB_TYPE then      
            dbms_sql.column_value(sd.sqlMetadata.cursorNumber, i, data.clob_value);
            if data.clob_value is not null and dbms_lob.getlength(data.clob_value) != 0 then
              -- try conversion to VARCHAR2
              begin
                data.varchar2_value := to_char(data.clob_value);
                sst_idx := put_string(ctx, data.varchar2_value);
                xutl_xlsb.put_CellIsst(stream, i-1, 0, sst_idx-1);
              exception
                when value_error then
                  -- stream CLOB content as an inline string, up to 32767 chars
                  xutl_xlsb.put_CellSt(stream, i-1, 0, lobValue => data.clob_value);
              end;
            end if;
            
          end case;
          
        end loop;
        
        sd.sqlMetadata.r_num := sd.sqlMetadata.r_num + 1;
        
        if rowIdx = MAX_ROW_NUMBER then
          if not sd.sqlMetadata.partitionBySize then
            -- force closing cursor
            nrows := 0;
          end if;
          exit;
        end if;
        
        exit when sd.sqlMetadata.r_num = partitionStop;
        
        -- fetch next row
        nrows := dbms_sql.fetch_rows(sd.sqlMetadata.cursorNumber);
          
      end loop;
      
      debug('end fetch');
      
      xutl_xlsb.put_simple_record(stream, 146);  -- BrtEndSheetData
      
      sheetRange := makeRange(sd.sqlMetadata.columnList(sd.sqlMetadata.visibleColumnSet.first).colRef
                            , 1
                            , sd.sqlMetadata.columnList(sd.sqlMetadata.visibleColumnSet.last).colRef
                            , rowIdx);
      
      -- autoFilter
      if sd.header.show and sd.header.autoFilter then
        if not sd.formatAsTable then
          sheet.filterRange := sheetRange;
          ctx.workbook.hasDefinedNames := true;
          xutl_xlsb.put_BeginAFilter(
            stream
          , firstRow    => sheetRange.start_ref.r - 1
          , firstCol    => sheetRange.start_ref.cn - 1
          , lastRow     => sheetRange.end_ref.r - 1
          , lastCol     => sheetRange.end_ref.cn - 1
          );
          xutl_xlsb.put_simple_record(stream, 162);  -- BrtEndAFilter
        end if;
      end if;
         
      -- new sheet
      sd.sqlMetadata.partitionId := sd.sqlMetadata.partitionId + 1;
      ctx.workbook.sheets.extend;
      sheet.sheetId := ctx.workbook.sheets.last;
      sheet.name := sd.sheetName;
      if sd.sqlMetadata.partitionBySize then
        sheet.name := replace(sheet.name, '${PNUM}', to_char(sd.sqlMetadata.partitionId));
        sheet.name := replace(sheet.name, '${PSTART}', to_char(partitionStart));
        sheet.name := replace(sheet.name, '${PSTOP}', to_char(sd.sqlMetadata.r_num));
      end if;
      
      -- check name validity
      if translate(sheet.name, '_\/*?:[]', '_') != sheet.name 
         or substr(sheet.name, 1, 1) = '''' 
         or substr(sheet.name, -1) = ''''
         or length(sheet.name) > 31 
      then
        error('Invalid sheet name: %s', sheet.name);
      end if;
      
      -- check name uniqueness
      if ctx.workbook.sheetMap.exists(sheet.name) then
        error('Duplicate sheet name: %s', sheet.name);
      end if;
      
      sheet.partName := 'xl/worksheets/sheet'||to_char(sheet.sheetId)||'.bin';
      sheet.tableParts := CT_TableParts();

      -- new sheet part
      part.name := sheet.partName;
      part.contentType := MT_WORKSHEET_BIN;
      part.rels := CT_Relationships();
      
      if sd.formatAsTable then
        tableId := addTable(ctx, sheetRange, sd.header.show, sd.header.autoFilter, sd.tableStyle, sd.columnMap, null, isEmpty);
        sheet.tableParts.extend;
        sheet.tableParts(sheet.tableParts.last) := tableId;
        
        -- table parts
        if sheet.tableParts.count != 0 then
          xutl_xlsb.put_simple_record(stream, 660, int2raw(sheet.tableParts.count)); -- BrtBeginListParts
          for i in 1 .. sheet.tableParts.count loop
            rId := addRelationship(part, RS_TABLE, ctx.workbook.tables(sheet.tableParts(i)).partName);
            xutl_xlsb.put_ListPart(stream, rId);  -- BrtListPart
          end loop;
          xutl_xlsb.put_simple_record(stream, 662);  -- BrtEndListParts
        end if;
      end if;
      
      xutl_xlsb.put_simple_record(stream, 130);  -- BrtEndSheet
      
      xutl_xlsb.flush_stream(stream);
      
      part.contentBin := stream.content;
      part.isBinary := true;
      
      -- add sheet to workbook
      ctx.workbook.sheets(sheet.sheetId) := sheet;
      ctx.workbook.sheetMap(sheet.name) := sheet.sheetId;
      
      -- add sheet part to package
      addPart(ctx, part);
    
    end if;
      
    if nrows = 0 then
      debug('close cursor');
      dbms_sql.close_cursor(sd.sqlMetadata.cursorNumber);
    end if;

  end;

  procedure createWorksheet (
    ctx         in out nocopy context_t
  , sheetIndex  in pls_integer
  )
  is
    sheetDefinition  sheet_definition_t;
    --columnFmt        varchar2(128);
    defaultFmt       varchar2(128);
    cellXf           CT_Xf;
    baseXfId         pls_integer := 0;
    columnId         pls_integer;
    columnName       varchar2(128);
    sheetColumn      sheet_column_t;
  begin
    sheetDefinition := ctx.sheetDefinitionMap(sheetIndex);
    prepareCursor(sheetDefinition.sqlMetadata);
    
    -- set column-level information
    for i in 1 .. sheetDefinition.sqlMetadata.columnList.count loop
      
      if not sheetDefinition.sqlMetadata.columnList(i).excluded then
        
        columnId := sheetDefinition.sqlMetadata.columnList(i).id; -- visible column ID
        columnName := sheetDefinition.sqlMetadata.columnList(i).name;
        sheetColumn := null;
        
        if sheetDefinition.columnMap.exists(columnId) then
          sheetColumn := sheetDefinition.columnMap(columnId);
        end if;
        
        if sheetColumn.header is null then
          sheetColumn.header := columnName;
          sheetDefinition.columnMap(columnId) := sheetColumn;
        end if;
        
        defaultFmt := case sheetDefinition.sqlMetadata.columnList(i).type
                      when dbms_sql.NUMBER_TYPE then coalesce(sheetDefinition.defaultFmts.numFmt, ctx.defaultFmts.numFmt, DEFAULT_NUM_FMT)
                      when dbms_sql.DATE_TYPE then coalesce(sheetDefinition.defaultFmts.dateFmt, ctx.defaultFmts.dateFmt, DEFAULT_DATE_FMT)
                      when dbms_sql.TIMESTAMP_TYPE then coalesce(sheetDefinition.defaultFmts.timestampFmt, ctx.defaultFmts.timestampFmt, DEFAULT_TIMESTAMP_FMT)
                      when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then coalesce(sheetDefinition.defaultFmts.timestampFmt, ctx.defaultFmts.timestampFmt, DEFAULT_TIMESTAMP_FMT)
                      end;
        
        sheetColumn.fmt := nvl(sheetColumn.fmt, defaultFmt);
        
        if sheetDefinition.columnLinkMap.exists(columnId) then
          baseXfId := ctx.workbook.styles.hlinkXfId;
        end if;
        
        cellXf := makeCellXf(ctx.workbook.styles, baseXfId, sheetColumn.fmt);
        sheetDefinition.sqlMetadata.columnList(i).xfId := putCellXf(ctx.workbook.styles, cellXf);
      
      end if;
      
    end loop;
    
    -- hyperlinks
    prepareHyperlinks(sheetDefinition);
    
    while dbms_sql.is_open(sheetDefinition.sqlMetadata.cursorNumber) loop
      case ctx.fileType
      when FILE_XLSX then
        createWorksheetImpl(ctx, sheetDefinition);
      when FILE_XLSB then
        createWorksheetBinImpl(ctx, sheetDefinition);
      end case;
    end loop;

  end;
  
  procedure createTable (
    ctx      in out nocopy context_t 
  , tableId  in pls_integer
  )
  is
    tab     CT_Table := ctx.workbook.tables(tableId);
    stream  stream_t := new_stream();
  begin
    stream_write(stream, '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="'||to_char(tab.id)||'" name="'||tab.name||'" displayName="'||tab.name||'" ref="'||tab.ref.expr||'"'||
                         case when not tab.showHeader then ' headerRowCount="0"' end ||
                         '>');
    if tab.showHeader and tab.autoFilter then
      stream_write(stream, '<autoFilter ref="'||tab.ref.expr||'"/>');
    end if;
    stream_write(stream, '<tableColumns count="'||to_char(tab.cols.count)||'">');
    for i in 1 .. tab.cols.count loop
      stream_write(stream, '<tableColumn id="'||tab.cols(i).id||'" name="'||dbms_xmlgen.convert(tab.cols(i).name)||'"/>');
    end loop;
    stream_write(stream, '</tableColumns>');
    stream_write(stream, '<tableStyleInfo' || 
                         case when tab.styleName is not null then ' name="'||tab.styleName||'"' end ||
                         ' showRowStripes="1"/>');
    stream_write(stream, '</table>');
    stream_flush(stream);
    addPart(ctx, tab.partName, MT_TABLE, stream.content);
  end;
  
  procedure createTableBin (
    ctx      in out nocopy context_t 
  , tableId  in pls_integer
  )
  is
    tab     CT_Table := ctx.workbook.tables(tableId);
    stream  xutl_xlsb.Stream_T := xutl_xlsb.new_stream();
  begin

    xutl_xlsb.put_BeginList(
      stream
    , tableId     => tab.id
    , name        => tab.name
    , displayName => tab.name
    , showHeader  => tab.showHeader
    , firstRow    => tab.ref.start_ref.r - 1
    , firstCol    => tab.ref.start_ref.cn - 1
    , lastRow     => tab.ref.end_ref.r - 1
    , lastCol     => tab.ref.end_ref.cn - 1
    );
    
    if tab.showHeader and tab.autoFilter then
      xutl_xlsb.put_BeginAFilter(
        stream
      , firstRow    => tab.ref.start_ref.r - 1
      , firstCol    => tab.ref.start_ref.cn - 1
      , lastRow     => tab.ref.end_ref.r - 1
      , lastCol     => tab.ref.end_ref.cn - 1
      );
      xutl_xlsb.put_simple_record(stream, 162);  -- BrtEndAFilter
    end if;
    
    xutl_xlsb.put_simple_record(stream, 345, int2raw(tab.cols.count));  -- BrtBeginListCols
    for i in 1 .. tab.cols.count loop
      xutl_xlsb.put_BeginListCol(stream, tab.cols(i).id, tab.cols(i).name); -- BrtBeginListCol
      xutl_xlsb.put_simple_record(stream, 348);  -- BrtEndListCol
    end loop;
    xutl_xlsb.put_simple_record(stream, 346);  -- BrtEndListCols
    
    xutl_xlsb.put_TableStyleClient(stream, tab.styleName);  -- BrtTableStyleClient
    
    xutl_xlsb.put_simple_record(stream, 344);  -- BrtEndList
    xutl_xlsb.flush_stream(stream);
    addPart(ctx, tab.partName, MT_TABLE_BIN, stream.content);
  end;
  
  procedure createWorkbook (
    ctx   in out nocopy context_t
  )
  is
    stream  stream_t;
    part    part_t;
  begin
    part.name := 'xl/workbook.xml';
    part.contentType := MT_WORKBOOK;
    part.rels := CT_Relationships();
    stream := new_stream();
    
    stream_write(stream, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    stream_write(stream, '<fileVersion appName="xl" lastEdited="2" lowestEdited="2"/>');
    stream_write(stream, '<sheets>');
    
    for i in 1 .. ctx.workbook.sheets.count loop
      -- add sheet relationships
      ctx.workbook.sheets(i).rId := addRelationship(part, RS_WORKSHEET, ctx.workbook.sheets(i).partName);
      stream_write(stream, '<sheet name="'||dbms_xmlgen.convert(ctx.workbook.sheets(i).name)||
                               '" sheetId="'||ctx.workbook.sheets(i).sheetId||
                               '" r:id="'||ctx.workbook.sheets(i).rId||'"/>');
    end loop;
    
    stream_write(stream, '</sheets>');
    
    if ctx.workbook.hasDefinedNames then
      stream_write(stream, '<definedNames>');
      for i in 1 .. ctx.workbook.sheets.count loop
        if ctx.workbook.sheets(i).filterRange.expr is not null then
          stream_write(stream, '<definedName name="_xlnm._FilterDatabase" localSheetId="'||to_char(i-1)||'" hidden="1">');
          stream_write(stream, dbms_xmlgen.convert('''' || replace(ctx.workbook.sheets(i).name, '''', '''''') || '''') || '!' || getRangeExpr(ctx.workbook.sheets(i).filterRange, true));
          stream_write(stream, '</definedName>');
        end if;
      end loop;      
      stream_write(stream, '</definedNames>');
    end if;
    
    -- set calculation engine version to max value
    stream_write(stream, '<calcPr calcId="999999"/>');
    
    stream_write(stream, '</workbook>');
    stream_flush(stream);
    
    part.content := stream.content;
    debug(xmltype(part.content).getstringval(1,2));
    
    addPart(ctx, part);
    
    createStylesheet(ctx, ctx.workbook.styles, 'xl/styles.xml');
    addRelationship(ctx, part.name, RS_STYLES, 'xl/styles.xml');
    
    createSharedStrings(ctx);
    addRelationship(ctx, part.name, RS_SHAREDSTRINGS, 'xl/sharedStrings.xml');
    
    for tableId in 1 .. ctx.workbook.tables.count loop
      createTable(ctx, tableId);
    end loop;
    
    -- add package-level relationship to workbook part
    addRelationship(ctx, null, RS_OFFICEDOCUMENT, part.name);
    
  end;

  procedure createWorkbookBin (
    ctx   in out nocopy context_t
  )
  is
    stream  xutl_xlsb.Stream_T := xutl_xlsb.new_stream();
    part    part_t;
    filterRange  range_t;
    links        xutl_xlsb.SupportingLinks_T;
  begin
    part.name := 'xl/workbook.bin';
    part.contentType := null;
    part.rels := CT_Relationships();
    
    xutl_xlsb.put_simple_record(stream, 131); -- BrtBeginBook
    xutl_xlsb.put_defaultBookViews(stream);
    xutl_xlsb.put_simple_record(stream, 143); -- BrtBeginBundleShs
    
    for i in 1 .. ctx.workbook.sheets.count loop
      -- add sheet relationships
      ctx.workbook.sheets(i).rId := addRelationship(part, RS_WORKSHEET, ctx.workbook.sheets(i).partName);
      xutl_xlsb.put_BundleSh(stream, ctx.workbook.sheets(i).sheetId, ctx.workbook.sheets(i).rId, ctx.workbook.sheets(i).name);
    end loop;
    
    xutl_xlsb.put_simple_record(stream, 144); -- BrtEndBundleShs
    
    if ctx.workbook.hasDefinedNames then
      
      xutl_xlsb.put_simple_record(stream, 353); -- BrtBeginExternals
      xutl_xlsb.put_simple_record(stream, 357); -- BrtSupSelf
      
      -- generate supporting links
      for i in 1 .. ctx.workbook.sheets.count loop
        if ctx.workbook.sheets(i).filterRange.expr is not null then
          ctx.workbook.sheets(i).filterXti := 
            xutl_xlsb.add_SupportingLink(links
                                        , 0    -- ref to BrtSupSelf record
                                        , i-1  -- bundleSh index
                                        );
        end if;
      end loop;
      xutl_xlsb.put_ExternSheet(stream, links);  -- BrtExternSheet
      
      xutl_xlsb.put_simple_record(stream, 354); -- BrtEndExternals
    
      for i in 1 .. ctx.workbook.sheets.count loop
        if ctx.workbook.sheets(i).filterRange.expr is not null then
          filterRange := ctx.workbook.sheets(i).filterRange;
          xutl_xlsb.put_FilterDatabase(stream
                                     , bundleShIndex => i-1
                                     , xti           => ctx.workbook.sheets(i).filterXti
                                     , firstRow      => filterRange.start_ref.r - 1
                                     , firstCol      => filterRange.start_ref.cn - 1
                                     , lastRow       => filterRange.end_ref.r - 1
                                     , lastCol       => filterRange.end_ref.cn - 1
                                     );
        end if;
      end loop;
      
    end if;
    
    xutl_xlsb.put_CalcProp(stream, 999999); -- BrtCalcProp
    
    xutl_xlsb.put_simple_record(stream, 132);  -- BrtEndBook
    xutl_xlsb.flush_stream(stream);
    part.contentBin := stream.content;
    part.isBinary := true;
    addPart(ctx, part);
    
    createStylesheetBin(ctx, ctx.workbook.styles, 'xl/styles.bin');
    addRelationship(ctx, part.name, RS_STYLES, 'xl/styles.bin');
    
    createSharedStringsBin(ctx);
    addRelationship(ctx, part.name, RS_SHAREDSTRINGS, 'xl/sharedStrings.bin');
    
    for tableId in 1 .. ctx.workbook.tables.count loop
      createTableBin(ctx, tableId);
    end loop;
    
    -- add package-level relationship to workbook part
    addRelationship(ctx, null, RS_OFFICEDOCUMENT, part.name);
    
  end;

  procedure createContentTypes (
    ctx   in out nocopy context_t
  )
  is
    stream  stream_t := new_stream();
  begin
    
    stream_write(stream, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    -- default extensions
    stream_write(stream, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    stream_write(stream, '<Default Extension="xml" ContentType="application/xml"/>');
    if ctx.fileType = FILE_XLSB then
      stream_write(stream, '<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>');
    end if;
    
    for i in 1 .. ctx.pck.parts.count loop
      if ctx.pck.parts(i).contentType is not null then
        stream_write(stream, '<Override PartName="/'||ctx.pck.parts(i).name||'" ContentType="'||ctx.pck.parts(i).contentType||'"/>');
      end if;
    end loop;    
    
    stream_write(stream, '</Types>');
    stream_flush(stream);
    
    debug(xmltype(stream.content).getstringval(1,2));
    
    addPart(ctx, '[Content_Types].xml', null, stream.content);
    
  end;
  
  function pack (
    part  in part_t
  )
  return blob
  is
    binaryContent  blob;
    binaryContentSize integer;
    gzContent      blob;
    gzSize         pls_integer;
    dt             timestamp := cast(sysdate as timestamp);
    output         blob;
    filename       raw(32767) := utl_i18n.string_to_raw(part.name, 'AL32UTF8');
   
    procedure write(bytes in raw)
    is
    begin
      dbms_lob.writeappend(output, utl_raw.length(bytes), bytes);
    end;
    procedure write_int(n in pls_integer, sz in pls_integer)
    is
    begin
      write(utl_raw.substr(utl_raw.cast_from_binary_integer(n, utl_raw.little_endian), 1, sz));
    end;

  begin
    
    binaryContent := case when part.isBinary then part.contentBin else xmlToBlob(part.content) end;
    binaryContentSize := dbms_lob.getlength(binaryContent);
    gzContent := utl_compress.lz_compress(binaryContent);
    gzSize := dbms_lob.getlength(gzContent);
    
    dbms_lob.freetemporary(binaryContent);
    
    dbms_lob.createtemporary(output, true);
    
    write('504B0304');  -- Local file header signature
    write('1400');      -- Version needed to extract (2.0)
    write('0008');      -- General purpose bit flag (bit11 = UTF-8 filename)
    write('0800');      -- Compression method (DEFLATE)
    
    -- File last modification time (MS-DOS format)
    write_int(
        extract(second from dt) / 2   -- bits 0-4
      + extract(minute from dt) * 32  -- bits 5-10
      + extract(hour from dt) * 2048  -- bits 11-15
    , 2
    );
    
    -- File last modification date (MS-DOS format)
    write_int(
        extract(day from dt) / 2                -- bits 0-4
      + extract(month from dt) * 32             -- bits 5-8
      + ( extract(year from dt) - 1980 ) * 512  -- bits 9-15
    , 2
    );
    
    write(dbms_lob.substr(gzContent, 4, gzSize - 7)); -- CRC32
    write_int(gzSize - 18, 4);                        -- Compressed size = sizeof(gzContent - header[10] - trailer[8])
    write_int(binaryContentSize, 4);                  -- Uncompressed size
    write_int(utl_raw.length(filename), 2);           -- File name length
    write('0000');                                    -- Extra field length
    write(filename);                                  -- File name

    -- Compressed data, copied from gzip content
    dbms_lob.copy(
      dest_lob    => output
    , src_lob     => gzContent
    , amount      => gzSize - 18
    , dest_offset => dbms_lob.getlength(output) + 1
    , src_offset  => 11                              -- gzip header size + 1
    );
    
    dbms_lob.freetemporary(gzContent);

    return output;

  end;
  
  procedure createPackage (
    pck  in out nocopy package_t
  )
  is
    zip    zip_t;
    entry  zip_entry_t; 
    pos    integer := 1;
    centralDirectorySize integer;

    procedure write(bytes in raw)
    is
    begin
      dbms_lob.writeappend(zip.content, utl_raw.length(bytes), bytes);
    end;
    procedure write_int(n in pls_integer, sz in pls_integer)
    is
    begin
      write(utl_raw.substr(utl_raw.cast_from_binary_integer(n, utl_raw.little_endian), 1, sz));
    end;

  begin
    dbms_lob.createtemporary(zip.content, true);
    zip.entries := zip_entry_list_t();
    
    zip.entries.extend(pck.parts.count);
    for i in 1 .. pck.parts.count loop
      entry.offset := pos;
      entry.filename := pck.parts(i).name;
      debug(entry.filename);
      entry.content := pack(pck.parts(i));
      dbms_lob.append(zip.content, entry.content);
      pos := pos + dbms_lob.getlength(entry.content);
      dbms_lob.freetemporary(entry.content);
      
      if pck.parts(i).isBinary then
        dbms_lob.freetemporary(pck.parts(i).contentBin);
      else
        dbms_lob.freetemporary(pck.parts(i).content);
      end if;
      
      zip.entries(i) := entry;
    end loop;
    
    -- Central directory file header
    for i in 1 .. zip.entries.count loop
      write('504B0102');
      write('1400');
      -- copy of local file header, from [Version needed to extract] to [File name length] : 24 bytes
      write(dbms_lob.substr(zip.content, 24, zip.entries(i).offset + 4));
      write('0000');                            -- Extra field length
      write('0000');                            -- File comment length
      write('0000');                            -- Disk number where file starts
      write('0000');                            -- Internal file attributes
      write('80000000');                        -- External file attributes
      write_int(zip.entries(i).offset - 1, 4);  -- Relative offset of local file header (0-based)
      write(utl_i18n.string_to_raw(zip.entries(i).filename, 'AL32UTF8'));  -- File name
    end loop;
    
    centralDirectorySize := dbms_lob.getlength(zip.content) - pos + 1;
    
    -- End of central directory record
    write('504B0506'); 	                 -- End of central directory signature = 0x06054b50
    write('0000');                       -- Number of this disk
    write('0000');                       -- Disk where central directory starts
    write_int(zip.entries.count, 2); 	   -- Number of central directory records on this disk
    write_int(zip.entries.count, 2); 	   -- Total number of central directory records
    write_int(centralDirectorySize, 4);  -- Size of central directory (bytes)
    write_int(pos - 1, 4); 	             -- Offset of start of central directory, relative to start of archive
    write('0000');                       -- Comment length
    
    pck.content := zip.content;
    
  end;

  function createContext (
    p_type  in pls_integer default FILE_XLSX 
  )
  return ctxHandle
  is
    ctxId  ctxHandle := nvl(ctx_cache.last, 0) + 1;
    ctx    context_t;
  begin
    ctx.fileType := nvl(p_type, FILE_XLSX);
    ctx.pck.parts := part_list_t();
    ctx.pck.rels := CT_Relationships();
    ctx.workbook := new_workbook();
    ctx_cache(ctxId) := ctx;
    return ctxId;
  end;
  
  procedure closeContext (
    p_ctxId  in ctxHandle 
  )
  is
  begin
    ctx_cache(p_ctxId).string_map.delete;
    ctx_cache(p_ctxId).string_list.delete;
    ctx_cache.delete(p_ctxId);
  end;

  function addSheetImpl (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle
  is
    sd        sheet_definition_t;
    local_rc  sys_refcursor := p_rc;
  begin
    sd.sheetName := p_sheetName;
    sd.tabColor := validateColor(p_tabColor);
    sd.formatAsTable := false;
    if p_paginate then
      sd.sqlMetadata.partitionBySize := true;
      sd.sqlMetadata.partitionSize := nvl(p_pageSize, MAX_ROW_NUMBER);
    end if;
    
    if p_query is not null then
      sd.sqlMetadata.queryString := p_query;
      sd.sqlMetadata.bindVariables := bind_variable_list_t();
    else
      sd.sqlMetadata.cursorNumber := dbms_sql.to_cursor_number(local_rc);
    end if;
    
    if p_sheetIndex is not null then
      if not ctx_cache(p_ctxId).sheetDefinitionMap.exists(p_sheetIndex) then
        sd.sheetIndex := p_sheetIndex;
      else
        error('Duplicate sheet index: %d', p_sheetIndex);
      end if;
    else
      sd.sheetIndex := nvl(ctx_cache(p_ctxId).sheetDefinitionMap.last, 0) + 1;
    end if;
    
    sd.sqlMetadata.excludeSet := parseIntList(p_excludeCols, ',');
    
    ctx_cache(p_ctxId).sheetDefinitionMap(sd.sheetIndex) := sd;
    ctx_cache(p_ctxId).sheetIndexMap(sd.sheetName) := sd.sheetIndex;
    
    return sd.sheetIndex;
  end;
  
  procedure addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  is
    sheetId  sheetHandle;
  begin
    sheetId := addSheetFromQuery(p_ctxId, p_sheetName, p_query, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_excludeCols);
  end;

  function addSheetFromQuery (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_query       in varchar2
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle
  is
    sheetId  sheetHandle;
  begin
    if p_query is null then
      error('Query string argument cannot be null');
    else
      sheetId := addSheetImpl(p_ctxId, p_sheetName, p_query, null, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_excludeCols);
    end if;
    return sheetId;
  end;

  procedure addSheetFromCursor (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  is
    sheetId  sheetHandle;
  begin
    sheetId := addSheetFromCursor(p_ctxId, p_sheetName, p_rc, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_excludeCols);
  end;

  function addSheetFromCursor (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_rc          in sys_refcursor
  , p_tabColor    in varchar2 default null
  , p_paginate    in boolean default false
  , p_pageSize    in pls_integer default null
  , p_sheetIndex  in pls_integer default null
  , p_excludeCols in varchar2 default null
  )
  return sheetHandle
  is
    sheetId  sheetHandle;
  begin
    if p_rc is null then
      error('Ref cursor argument cannot be null');
    else
      sheetId := addSheetImpl(p_ctxId, p_sheetName, null, p_rc, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_excludeCols);
    end if;
    return sheetId;
  end;

  -- to be deprecated
  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  )
  is
  begin
    setHeader(p_ctxId, ctx_cache(p_ctxId).sheetIndexMap(p_sheetName), p_style, p_frozen, p_autofilter);
  end;

  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  )
  is
    sheetHeader  sheet_header_t;
  begin
    sheetHeader.show := true;
    sheetHeader.xfId := p_style;
    sheetHeader.isFrozen := p_frozen;
    sheetHeader.autoFilter := p_autoFilter;
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).header := sheetHeader;
  end;

  -- to be deprecated
  procedure setTableFormat (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_style      in varchar2 default null
  )
  is
  begin
    setTableFormat(p_ctxId, ctx_cache(p_ctxId).sheetIndexMap(p_sheetName), p_style);
  end;

  procedure setTableFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_style    in varchar2 default null
  )
  is
  begin
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).formatAsTable := true;
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).tableStyle := p_style;
  end;

  procedure setDateFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).defaultFmts.dateFmt := p_format;
  end;

  procedure setDateFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).defaultFmts.dateFmt := p_format;
  end;

  procedure setNumFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).defaultFmts.numFmt := p_format;
  end;

  procedure setNumFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).defaultFmts.numFmt := p_format;
  end;

  procedure setTimestampFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).defaultFmts.timestampFmt := p_format;
  end;

  procedure setTimestampFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).defaultFmts.timestampFmt := p_format;
  end;
  
  procedure setColumnFormat (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_columnId  in pls_integer
  , p_format    in varchar2 default null
  , p_header    in varchar2 default null
  , p_width     in number default null
  )
  is
    sheetColumn  sheet_column_t;
  begin
    sheetColumn.fmt := p_format;
    sheetColumn.header := p_header;
    sheetColumn.width := p_width;
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).columnMap(p_columnId) := sheetColumn;
  end;

  procedure setColumnHlink (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_columnId  in pls_integer
  , p_target    in varchar2 default null
  --, p_tooltip   in varchar2 default null
  )
  is
  begin   
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetId).columnLinkMap(p_columnId) := p_target;
    ctx_cache(p_ctxId).workbook.styles.hasHlink := true;
  end;

  procedure setBindVariableImpl (
    p_ctxId       in ctxHandle
  , p_sheetIndex  in pls_integer
  , p_bindName    in varchar2
  , p_bindValue   in anydata
  )
  is
    sheetDefinition  sheet_definition_t;
    bindVar          bind_variable_t;
    collIdx          pls_integer;
  begin
    bindVar.name := p_bindName;
    bindVar.value := p_bindValue;
    sheetDefinition := ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetIndex);
    
    sheetDefinition.sqlMetadata.bindVariables.extend;
    collIdx := sheetDefinition.sqlMetadata.bindVariables.last;
    sheetDefinition.sqlMetadata.bindVariables(collIdx) := bindVar;
    
    ctx_cache(p_ctxId).sheetDefinitionMap(p_sheetIndex) := sheetDefinition;
    
  end;

  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in number
  )
  is
  begin
    setBindVariableImpl(p_ctxId, ctx_cache(p_ctxId).sheetIndexMap(p_sheetName), p_bindName, anydata.ConvertNumber(p_bindValue));
  end;

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in number
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertNumber(p_bindValue));
  end;
  
  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  )
  is
  begin
    setBindVariableImpl(p_ctxId, ctx_cache(p_ctxId).sheetIndexMap(p_sheetName), p_bindName, anydata.ConvertVarchar2(p_bindValue));
  end;
  
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertVarchar2(p_bindValue));
  end;

  -- to be deprecated
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in date
  )
  is
  begin
    setBindVariableImpl(p_ctxId, ctx_cache(p_ctxId).sheetIndexMap(p_sheetName), p_bindName, anydata.ConvertDate(p_bindValue));
  end;

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in date
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertDate(p_bindValue));
  end;
  
$if NOT $$no_crypto OR $$no_crypto IS NULL $then
  procedure setEncryption (
    p_ctxId       in ctxHandle
  , p_password    in varchar2
  , p_compatible  in pls_integer default OFFICE2007SP2
  )
  is
    encInfo  encryption_info_t;
  begin
    case p_compatible
    when OFFICE2007SP1 then
      encInfo.version := '3.2';
      encInfo.cipherName := 'AES128';
      encInfo.hashFuncName := 'SHA1';
    when OFFICE2007SP2 then
      encInfo.version := '4.2';
      encInfo.cipherName := 'AES128';
      encInfo.hashFuncName := 'SHA1';
    when OFFICE2010 then
      encInfo.version := '4.4';
      encInfo.cipherName := 'AES128';
      encInfo.hashFuncName := 'SHA1';
    when OFFICE2013 then
      encInfo.version := '4.4';
      encInfo.cipherName := 'AES256';
      encInfo.hashFuncName := 'SHA512';
    when OFFICE2016 then
      encInfo.version := '4.4';
      encInfo.cipherName := 'AES256';
      encInfo.hashFuncName := 'SHA512';
    else
      error('Invalid compatible parameter : %d', p_compatible);
    end case;
    
    encInfo.password := p_password;
    
    ctx_cache(p_ctxId).encryptionInfo := encInfo;
      
  end;
$end

  function getFileContent (
    p_ctxId  in ctxHandle
  )
  return blob
  is
    ctx         context_t := ctx_cache(p_ctxId);
    sheetIndex  pls_integer;
    output      blob;
  begin
    -- shared styles
    addDefaultStyles(ctx.workbook.styles);
  
    -- worksheets
    sheetIndex := ctx.sheetDefinitionMap.first;
    while sheetIndex is not null loop
      createWorksheet(ctx, sheetIndex);
      sheetIndex := ctx.sheetDefinitionMap.next(sheetIndex);
    end loop;
    
    -- workbook
    case ctx.fileType
    when FILE_XLSX then
      createWorkbook(ctx);
    when FILE_XLSB then
      createWorkbookBin(ctx);
    end case;
    
    createContentTypes(ctx);
    createRels(ctx);
    
    debug('start create package');  
    createPackage(ctx.pck);  
    debug('end create package');
    
$if NOT $$no_crypto OR $$no_crypto IS NULL $then
    if ctx.encryptionInfo.version is not null then
      output := xutl_offcrypto.encrypt_package(
                  p_package  => ctx.pck.content
                , p_password => ctx.encryptionInfo.password
                , p_version  => ctx.encryptionInfo.version
                , p_cipher   => ctx.encryptionInfo.cipherName
                , p_hash     => ctx.encryptionInfo.hashFuncName
                );
      dbms_lob.freetemporary(ctx.pck.content);
    else    
$end
      output := ctx.pck.content;
$if NOT $$no_crypto OR $$no_crypto IS NULL $then
    end if;
$end
    
    return output;
    
  end;
  
  procedure createFile (
    p_ctxId      in ctxHandle
  , p_directory  in varchar2
  , p_filename   in varchar2
  )
  is
    fileContent  blob := getFileContent(p_ctxId);
  begin
    writeBlobToFile(p_directory, p_filename, fileContent);
    dbms_lob.freetemporary(fileContent);
  end;

begin
  
  init;

end ExcelGen;
/
