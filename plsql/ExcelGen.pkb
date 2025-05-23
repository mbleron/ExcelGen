create or replace package body ExcelGen is
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
    Marc Bleron       2024-12-07     Fixed sheetIndexMap access using upper-ized sheet name
                                     Do not call ExcelFormula.setCurrentSheet for a pageable sheet
                                     Fixed insertFirst/Last calls
    Marc Bleron       2025-01-31     Fixed corrupted sheet references in formulas after a pageable sheet
                                     More fixes related to pageable sheets and virtual columns
    Marc Bleron       2025-02-08     Image support
    Marc Bleron       2025-05-08     Sheet background
====================================================================================== */

  VERSION_NUMBER     constant varchar2(16) := '4.5.0';

  -- OPC part MIME types
  --MT_COMMENTS        constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
  MT_STYLES                constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  MT_WORKBOOK              constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  MT_WORKSHEET             constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  MT_SHAREDSTRINGS         constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
  MT_TABLE                 constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
  
  MT_CORE                  constant varchar2(256) := 'application/vnd.openxmlformats-package.core-properties+xml';
  MT_DRAWING               constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.drawing+xml';
  MT_METADATA              constant varchar2(256) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
  MT_RDRICHVALUESTRUCTURE  constant varchar2(256) := 'application/vnd.ms-excel.rdrichvaluestructure+xml';
  MT_RICHVALUEREL          constant varchar2(256) := 'application/vnd.ms-excel.richvaluerel+xml';
  MT_RDRICHVALUE           constant varchar2(256) := 'application/vnd.ms-excel.rdrichvalue+xml';
  
  MT_PNG                   constant varchar2(256) := 'image/png';
  MT_JPEG                  constant varchar2(256) := 'image/jpeg';
  MT_GIF                   constant varchar2(256) := 'image/gif';
  MT_SVG                   constant varchar2(256) := 'image/svg+xml';
  
  -- Binary MIME types
  MT_STYLES_BIN            constant varchar2(256) := 'application/vnd.ms-excel.styles';
  MT_WORKSHEET_BIN         constant varchar2(256) := 'application/vnd.ms-excel.worksheet';
  MT_SHAREDSTRINGS_BIN     constant varchar2(256) := 'application/vnd.ms-excel.sharedStrings';
  MT_TABLE_BIN             constant varchar2(256) := 'application/vnd.ms-excel.table';
  MT_METADATA_BIN          constant varchar2(256) := 'application/vnd.ms-excel.sheetMetadata';
  
  -- Relationship types
  --RS_COMMENTS        constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
  RS_OFFICEDOCUMENT        constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RS_WORKSHEET             constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  RS_STYLES                constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  RS_SHAREDSTRINGS         constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
  RS_TABLE                 constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
  RS_CORE                  constant varchar2(256) := 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
  RS_DRAWING               constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
  RS_IMAGE                 constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  RS_METADATA              constant varchar2(256) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
  RS_RDRICHVALUESTRUCTURE  constant varchar2(256) := 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure';
  RS_RICHVALUEREL          constant varchar2(256) := 'http://schemas.microsoft.com/office/2022/10/relationships/richValueRel';
  RS_RDRICHVALUE           constant varchar2(256) := 'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue';

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
  
  DEFAULT_COL_WIDTH      constant number := 10.71;
  DEFAULT_DATE_FMT       constant varchar2(32) := 'dd/mm/yyyy hh:mm:ss';
  DEFAULT_TIMESTAMP_FMT  constant varchar2(32) := 'dd/mm/yyyy hh:mm:ss.000';
  DEFAULT_NUM_FMT        constant varchar2(32) := null; 
  NLS_PARAM_STRING       constant varchar2(32) := 'nls_numeric_characters=''. ''';
  
  -- supertypes
  ST_NUMBER              constant pls_integer := 0;
  ST_STRING              constant pls_integer := 1;
  ST_DATETIME            constant pls_integer := 2;
  ST_LOB                 constant pls_integer := 3;
  ST_VARIANT             constant pls_integer := 4;
  ST_RICHTEXT            constant pls_integer := 5;
  ST_FORMULA             constant pls_integer := 6;
  ST_IMAGE               constant pls_integer := 7;

  buffer_too_small       exception;
  pragma exception_init (buffer_too_small, -19011);
  
  xml_parse_exception    exception;
  pragma exception_init (xml_parse_exception, -31011);

  type stream_t is record (
    content   clob
  , buf       varchar2(32767)
  , buf_sz    pls_integer
  );

  type data_t is record (
    st              pls_integer  -- supertype
  , db_type         pls_integer
  , varchar2_value  varchar2(32767)
  , char_value      char(32767)
  , number_value    number
  , date_value      date
  , ts_value        timestamp
  , tstz_value      timestamp with time zone
  , clob_value      clob
  , anydata_value   anydata
  , xml_value       xmltype
  , blob_value      blob
  );
  
  type data_map_t is table of data_t index by pls_integer;
  
  type cell_ref_t is record (value varchar2(10), c varchar2(3), cn pls_integer, r pls_integer); 
  type range_t is record (expr varchar2(32), start_ref cell_ref_t, end_ref cell_ref_t, is_null boolean := true);
  
  type intList_t is table of pls_integer;
  type intSet_t is table of pls_integer index by pls_integer;
  type stringSet_t is table of pls_integer index by varchar2(512);
  type valueSet_t is record (integers intSet_t, strings stringSet_t);
  
  type link_token_map_t is table of varchar2(8) index by pls_integer;
  type link_t is record (target varchar2(2048), tooltip varchar2(256), tokens link_token_map_t, fmla varchar2(8192));
  --type link_rel_map_t is table of varchar2(256) index by varchar2(2048);
  type link_map_t is table of varchar2(2048) index by pls_integer;

  type column_ref_list_t is table of varchar2(3);
  
  type formula_t is record (expr varchar2(32767), shared boolean, sharedIdx pls_integer, hasRef boolean, refStyle pls_integer);
  
  type column_t is record (
    name     varchar2(128)
  , type     pls_integer
  , scale    pls_integer
  , id       pls_integer
  , dbId     pls_integer
  , colRef   varchar2(3)
  , colNum   pls_integer
  , xfId     pls_integer := 0
  , hyperlink  boolean := false
  , linkTokens  intList_t
  , supertype pls_integer
  , excluded  boolean := false
  , fmla      formula_t
  , virtual   boolean := false
  );
    
  type column_list_t is table of column_t;
  type column_map_t is table of pls_integer index by varchar2(128);
  type column_set_t is table of varchar2(3) index by pls_integer;

  type string_map_t is table of pls_integer index by varchar2(32767);
  type string_t is record (value varchar2(32767), isRichText boolean := false);
  type string_list_t is table of string_t;
  type richText_cache_t is table of ExcelTypes.CT_RichText index by pls_integer;
  
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
  
  type CT_Dxf is record (
    style    ExcelTypes.CT_Style
  , content  varchar2(32767)
  );
  
  type CT_DxfMap is table of pls_integer index by varchar2(32767);
  type CT_Dxfs is table of CT_Dxf index by pls_integer;
  
  type CT_CellXfMap is table of pls_integer index by varchar2(32767);
  type CT_CellXfs is table of CT_Xf index by pls_integer;
  
  type CT_CellStyle is record (
    name       varchar2(256)
  , xfId       pls_integer
  , builtinId  pls_integer
  );
  
  type CT_CellStyles is table of CT_CellStyle;
  
  type CT_Stylesheet is record (
    numFmtIdSeq     pls_integer
  , numFmtMap       CT_NumFmtMap
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
  , dxfMap          CT_DxfMap
  , dxfs            CT_Dxfs
  , hasHlink        boolean := false
  , hlinkXfId       pls_integer
  );
  
  type CT_TableColumn is record (
    id    pls_integer
  , name  varchar2(256)
  );
  
  type CT_TableColumns is table of CT_TableColumn;
  
  type CT_Table is record (
    id                 pls_integer
  , name               varchar2(256)
  , ref                range_t
  , cols               CT_TableColumns
  , showHeader         boolean
  , autoFilter         boolean
  , styleName          varchar2(64)
  , partName           varchar2(256)
  , showFirstColumn    boolean
  , showLastColumn     boolean
  , showRowStripes     boolean
  , showColumnStripes  boolean
  );
  
  type CT_Tables is table of CT_Table index by pls_integer;
  
  type CT_TableParts is table of pls_integer;
  
  type CT_Sheet is record (
    name         varchar2(128)
  , sheetId      pls_integer
  , rId          varchar2(256)
  , state        pls_integer
  , partName     varchar2(256)
  , tableParts   CT_TableParts
  );
  
  type CT_Sheets is table of CT_Sheet;
  type CT_SheetMap is table of pls_integer index by varchar2(128);
  
  type CT_Workbook is record (
    sheets      CT_Sheets
  , sheetMap    CT_SheetMap
  , styles      CT_Stylesheet
  , tables      CT_Tables
  , firstSheet  pls_integer -- first visible sheet idx
  , refStyle    pls_integer
  , partName    varchar2(256)
  );
  
  type bind_variable_t is record (
    name   varchar2(30)
  , value  anydata
  );
  
  type bind_variable_list_t is table of bind_variable_t;
  
  type virtualColumn_t is record (
    col    column_t
  , pos    pls_integer
  , after  boolean
  );
  
  type virtualColumnList_t is table of virtualColumn_t;

  -- doubly linked list of columns
  type dbLinkedNode_t is record (id pls_integer, data column_t, prev pls_integer, next pls_integer);
  type dbLinkedNodeHeap_t is table of dbLinkedNode_t index by pls_integer;
  type dbLinkedList_t is record (first pls_integer, last pls_integer, heap dbLinkedNodeHeap_t);
  
  type sql_metadata_t is record (
    queryString       clob
  , cursorNumber      integer
  , bindVariables     bind_variable_list_t
  , columnList        column_list_t
  , columnMap         column_map_t
  , virtualColumns    virtualColumnList_t
  , visibleColumnSet  column_set_t
  , excludeSet        valueSet_t
  , partitionBySize   boolean := false
  , partitionSize     pls_integer
  , partitionId       pls_integer
  , r_num             pls_integer
  , maxRows           integer
  );
  
  type table_column_t is record (
    name        varchar2(1024)
  , xfId        pls_integer
  , headerXfId  pls_integer
  , dvRule      ExcelTypes.CT_DataValidation
  , cfRules     ExcelTypes.CT_CfRules
  );
  
  type table_column_map_t is table of table_column_t index by pls_integer;
  
  type table_header_t is record (
    show        boolean
  , isFrozen    boolean
  , autoFilter  boolean
  );
  
  type colProperties_t is record (
    xfId    pls_integer
  , width   number
  );
  
  type colProperties_map_t is table of colProperties_t index by pls_integer;
  
  type rowProperties_t is record (
    xfId    pls_integer
  , height  number
  );
  
  type rowProperties_map_t is table of rowProperties_t index by pls_integer;
  
  type cell_t is record (
    r            pls_integer
  , c            varchar2(3)
  , cn           pls_integer
  , xfId         pls_integer
  , v            data_t
  , f            formula_t
  , hyperlink    boolean
  , isTableCell  boolean := false
  );
  
  type cellList_t is table of cell_t index by pls_integer;  
  type row_t is record (id pls_integer, props rowProperties_t, cells cellList_t);
  type rowList_t is table of row_t index by pls_integer;
  type sheetData_t is record (rows rowList_t, hasCells boolean);
  
  type anchorRef_t is record (tableId pls_integer, anchorPosition pls_integer, rowOffset pls_integer, colOffset pls_integer);
  
  type floatingCell_t is record (
    data       data_t
  , xfId       pls_integer
  , anchorRef  anchorRef_t
  , fmla       formula_t
  , hyperlink  boolean
  );
  
  type floatingCellList_t is table of floatingCell_t;
  
  type cellSpan_t is record (anchorRef anchorRef_t, rowSpan pls_integer, colSpan pls_integer);
  type cellSpanList_t is table of cellSpan_t;
  
  type cellRange_t is record (name varchar2(256), span cellSpan_t, xfId pls_integer, outsideBorders boolean);
  type cellRangeList_t is table of cellRange_t;
  
  type table_t is record (
    id                 pls_integer
  , header             table_header_t
  , formatAsTable      boolean
  , tableName          varchar2(1024)
  , tableStyle         varchar2(32)
  , sqlMetadata        sql_metadata_t
  , columnLinkMap      link_map_t
  , range              range_t
  , isEmpty            boolean
  , columnMap          table_column_map_t
  , rowMap             rowProperties_map_t
  , showFirstColumn    boolean
  , showLastColumn     boolean
  , showRowStripes     boolean
  , showColumnStripes  boolean
  , anchorRef          anchorRef_t
  , cfRules            ExcelTypes.CT_CfRules
  , rowProps           rowProperties_t
  );
  
  type tableTreeNode_t is record (nodeId pls_integer, children intList_t);
  type tableTree_t is table of tableTreeNode_t index by pls_integer;
  type tableForest_t is record (roots intList_t, t tableTree_t);
  type tableList_t is table of table_t;
  
  type sharedFmlaRef_t is record (columnId pls_integer, tableId pls_integer);
  type sharedFmlaMap_t is table of sharedFmlaRef_t index by pls_integer;
  
  type CT_Point2D is record (x varchar2(128), y varchar2(128));
  type CT_PositiveSize2D is record (cx varchar2(128), cy varchar2(128));
  type CT_Marker is record (col pls_integer, colOff varchar2(128), "row" pls_integer, rowOff varchar2(128)); 
  type CT_DrawingAnchor is record (
    id          pls_integer
  , anchorType  pls_integer
  , "from"      CT_Marker
  , "to"        CT_Marker
  , pos         CT_Point2D
  , ext         CT_PositiveSize2D
  , editAs      pls_integer
  , content     blob
  );
  type CT_Drawing is table of CT_DrawingAnchor;
  
  type sheet_definition_t is record (
    sheetName            varchar2(128)
  , sheetIndex           pls_integer
  , tabColor             varchar2(8)
  , state                pls_integer
  , defaultFmts          defaultFmts_t
  , defaultXfId          pls_integer
  , columnMap            colProperties_map_t
  , hasCustomColProps    boolean
  --, columnLinkMap   link_map_t
  , tableList            tableList_t
  , tableForest          tableForest_t
  , data                 sheetData_t
  , pageable             boolean := false
  , streamable           boolean
  , done                 boolean
  , hasProps             boolean
  , activePaneAnchorRef  cell_ref_t
  , showGridLines        boolean
  , showRowColHeaders    boolean
  , defaultRowHeight     number
  , mergedCells          cellSpanList_t
  , floatingCells        floatingCellList_t
  , cellRanges           cellRangeList_t
  , sharedFmlaSeq        pls_integer
  , sharedFmlaMap        sharedFmlaMap_t
  , dvRules              ExcelTypes.CT_DataValidations
  , cfRules              ExcelTypes.CT_CfRules
  , drawing              CT_Drawing
  , background           blob
  );
  
  type sheet_definition_map_t is table of sheet_definition_t index by pls_integer;
    
  type encryption_info_t is record (
    version       varchar2(3)
  , cipherName    varchar2(16)
  , hashFuncName  varchar2(16)
  , password      varchar2(512)
  );
  
  type coreProperties_t is record (
    creator      varchar2(256)
  , description  varchar2(4000)
  , subject      varchar2(4000)
  , title        varchar2(4000)  
  );
    
  type image_t is record (
    partName  varchar2(256)
  , name      varchar2(256)
  , mimeType  varchar2(256)
  , width     pls_integer
  , height    pls_integer
  );
  
  type imageList_t is table of image_t;
  
  -- Custom hash map to hold image part names indexed by their MD5 checksums
  -- Key size is 32 = 16 (MD5 output) * 2 bytes (hex string)
  type imageHashMap_t is table of imageList_t index by varchar2(32);
  
  type imagePartMap_t is table of varchar2(256) index by pls_integer; --partName indexed by valueMetadataId
  type imageValueMetadataMap_t is table of pls_integer index by varchar2(256); -- valueMetadataId indexed by partName
  
  type imageMetadata_t is record (vmMap imageValueMetadataMap_t, partMap imagePartMap_t);
  
  type defaultExtensionMap_t is table of varchar2(16) index by varchar2(256);
  
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
  , defaultXfId          pls_integer
  , encryptionInfo       encryption_info_t
  , fileType             pls_integer
  , rt_cache             richText_cache_t
  , coreProperties       coreProperties_t
  , names                ExcelTypes.CT_DefinedNames
  , nameMap              ExcelTypes.CT_DefinedNameMap
  , tableNameSeq         pls_integer := 0
  , imageNameSeq         pls_integer := 0
  , imageMetadata        imageMetadata_t
  , imageHashMap         imageHashMap_t
  , extensions           defaultExtensionMap_t
  , rvStructures         ExcelTypes.CT_RichValueStructures
  , drawingIdSeq         pls_integer := 0
  );
  
  type context_cache_t is table of context_t index by pls_integer;
  
  type crc_table_t is table of simple_integer;
  
  ctx_cache       context_cache_t;
  currentCtx      context_t;
  currentCtxId    pls_integer := -1;
  
  debug_enabled   boolean := false;
  hash_available  boolean;
  crc_table       crc_table_t;
  
  function getProductName return varchar2
  is
  begin
    return 'EXCELGEN-' || VERSION_NUMBER;
  end;

  procedure loadContext (ctxId in pls_integer)
  is
  begin
    if ctxId != currentCtxId then
      ctx_cache(currentCtxId) := currentCtx;
      currentCtxId := ctxId;
      currentCtx := ctx_cache(currentCtxId);
    end if;
  end;
  
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

  procedure assertPositive (
    val      in number
  , message  in varchar2
  )
  is
  begin
    if not val > 0 then
      error(message);
    end if;
  end;
  
  function hashMD5 (input in raw) return raw 
  is
    output  raw(16);
  begin
    $IF DBMS_DB_VERSION.VERSION >= 12  $THEN
      select standard_hash(input, 'MD5') into output from dual;
    $ELSE
      execute immediate 'begin :1 := dbms_crypto.Hash(:2, dbms_crypto.HASH_MD5); end;' using out output, input;
    $END
    return output;
  end;

  procedure checkHashFunction is
    dummy raw(16);
  begin
    dummy := hashMD5('00');
    hash_available := true;
  exception
    when others then
      debug('DBMS_CRYPTO not available');
      hash_available := false;
  end;
  
  function xor (x in simple_integer, y in simple_integer) return simple_integer 
  is
  begin
    return x + y - 2 * bitand(x,y);
  end;
  
  procedure makeCRC32LookupTable is
    c  simple_integer := 0;
    b  simple_integer := utl_raw.cast_to_binary_integer('EDB88320');
    -- 1 bit right-shift
    function rsh1 (x in simple_integer) return simple_integer
    is
    begin
      if x < 0 then
        -- unset sign bit, floor-divide as a positive int, and reset shifted bit
        return floor(bitand(x, 2147483647)/2) + 1073741824;
      else
        return floor(x/2);
      end if;
    end;
  begin
    crc_table := crc_table_t();
    crc_table.extend(256);
    for n in 0 .. 255 loop 
      c := n;
      for k in 0 .. 7 loop
        if bitand(c, 1) != 0 then
          c := xor(b, rsh1(c));
        else
          c := rsh1(c);
        end if;
      end loop;
      crc_table(n+1) := c;
    end loop;    
  end;
  
  function adler32 (f in blob) return raw 
  is  
    CHUNK_SIZE  constant pls_integer := 32767;
    len         pls_integer;
    b           simple_integer := 0;
    buf         raw(32767);
    nchunks     integer := ceil(dbms_lob.getlength(f)/CHUNK_SIZE);
    s1          simple_integer := 1;
    s2          simple_integer := 0;
    n           simple_integer := 0;
    off         simple_integer := 0;
  begin
    for k in 0 .. nchunks - 1 loop
      buf := dbms_lob.substr(f, CHUNK_SIZE, k*CHUNK_SIZE+1);
      len := utl_raw.length(buf);
      off := 1;
      while len > 0 loop
        n := 3800;
        if n > len then
          n := len;
        end if;
        len := len - n;
        while n > 0 loop
          b := to_number(utl_raw.substr(buf, off, 1),'XX');
          s1 := s1 + b;
          s2 := s2 + s1;
          n := n - 1;
          off := off + 1;
        end loop;
        s1 := mod(s1, 65521);
        s2 := mod(s2, 65521);
      end loop;
    end loop;
    return utl_raw.cast_from_binary_integer(s2*65536 + s1);
  end;
  
  function crc32 (f in blob, offset in integer default null, amount in integer default null) return raw 
  is  
    startOffset  integer := nvl(offset,1);
    endOffset    integer := startOffset + nvl(amount, dbms_lob.getlength(f)) - 1;
    c    simple_integer := 0;
    i    simple_integer := utl_raw.cast_to_binary_integer('FFFFFFFF');
    b    simple_integer := 0;
    -- 8-bit right shift
    function rsh8 (x in simple_integer) return simple_integer
    is
    begin
      if x < 0 then
        -- unset sign bit, floor-divide as a positive int, and reset shifted bit
        return floor(bitand(x, 2147483647)/256) + 8388608;
      else
        return floor(x/256);
      end if;
    end;
  begin
    c := i;
    for n in startOffset .. endOffset loop
      b := to_number(dbms_lob.substr(f, 1, n),'XX');
      c := xor(crc_table(bitand(xor(c,b),255) + 1), rsh8(c));
    end loop;
    return utl_raw.cast_from_binary_integer(xor(c,i));
  end;  
  
  procedure init
  is  
  begin
    checkHashFunction;
    makeCRC32LookupTable;
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
  return intSet_t
  is
    i       pls_integer;
    token   varchar2(256);
    p1      pls_integer := 1;
    p2      pls_integer;
    output  intSet_t;
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

  function parseValueList (
    input  in varchar2
  )
  return valueSet_t
  is
  
    vals  valueSet_t;
    p1    pls_integer := 1;
    p2    pls_integer;
    token varchar2(512);
    c     varchar2(1 char);
    
    i     pls_integer;
    s     token%type;
    
    procedure skipws is
    begin
      while substr(input, p1, 1) = ' ' loop
        p1 := p1 + 1;
      end loop;
    end;
    
  begin
    
    if input is not null then
    
      loop
      
        skipws;
        
        if substr(input, p1, 1) = '"' then
          
          p1 := p1 + 1;
          p2 := instr(input, '"', p1); -- terminating quote
          if p2 = 0 then
            raise_application_error(-20000, utl_lms.format_message('Missing terminating quote'));
          else
            token := substr(input, p1, p2 - p1);
            p1 := p2 + 1;
            if token is not null then
              vals.strings(token) := 1;
            end if;
          end if;
          skipws;
          c := substr(input, p1, 1);
          if c = ',' then
            p1 := p1 + 1;
          elsif c is null then
            exit;
          else
            raise_application_error(-20000, utl_lms.format_message('Unexpected character at position %d: ''%s''', p1, c));
          end if;
          
        else
          
          p2 := instr(input, ',', p1);
          if p2 = 0 then
            token := rtrim(substr(input, p1));
            if token is not null then
              vals.integers(to_number(token)) := 1;
            end if;
            exit;
          else
            token := rtrim(substr(input, p1, p2 - p1));
            if token is not null then
              vals.integers(to_number(token)) := 1;
            end if;
            p1 := p2 + 1;
          end if;
          
        end if;
      
      end loop;
      
      if debug_enabled then
        s := vals.strings.first;
        while s is not null loop
          debug(s);
          s := vals.strings.next(s);
        end loop;
        
        i := vals.integers.first;
        while i is not null loop
          debug(i);
          i := vals.integers.next(i);
        end loop;
      end if;
    
    end if;

    return vals;
      
  end;

  procedure insertAfter (t in out nocopy dbLinkedList_t, nodeId in pls_integer, newNodeId in pls_integer) is
  begin
    t.heap(newNodeId).prev := nodeId;
    if t.heap(nodeId).next is null then
      t.heap(newNodeId).next := null;
      t.last := newNodeId;
    else
      t.heap(newNodeId).next := t.heap(nodeId).next;
      t.heap(t.heap(nodeId).next).prev := newNodeId;
    end if;
    t.heap(nodeId).next := newNodeId;
  end;
  
  procedure insertBefore (t in out nocopy dbLinkedList_t, nodeId in pls_integer, newNodeId in pls_integer) is
  begin
    t.heap(newNodeId).next := nodeId;
    if t.heap(nodeId).prev is null then
      t.heap(newNodeId).prev := null;
      t.first := newNodeId;
    else
      t.heap(newNodeId).prev := t.heap(nodeId).prev;
      t.heap(t.heap(nodeId).prev).next := newNodeId;
    end if;
    t.heap(nodeId).prev := newNodeId;
  end;
  
  procedure insertFirst (t in out nocopy dbLinkedList_t, nodeId in pls_integer) is
    nextNodeId  pls_integer;
  begin
    if t.first is null then
      t.first := nodeId;
      t.last := nodeId;
      t.heap(nodeId).prev := null;
      t.heap(nodeId).next := null;
    else
      nextNodeId := t.first;
      insertBefore(t, nextNodeId, nodeId);
    end if;
  end;

  procedure insertLast (t in out nocopy dbLinkedList_t, nodeId in pls_integer) is
    prevNodeId  pls_integer;
  begin
    if t.last is null then
      insertFirst(t, nodeId);
    else
      prevNodeId := t.last;
      insertAfter(t, prevNodeId, nodeId);
    end if;
  end;

  function getDefaultFormat (
    ctx     in context_t
  , sd      in sheet_definition_t
  , dbType  in pls_integer
  )
  return varchar2
  is
  begin
    return case dbType
           when dbms_sql.NUMBER_TYPE then coalesce(sd.defaultFmts.numFmt, ctx.defaultFmts.numFmt, DEFAULT_NUM_FMT)
           when dbms_sql.DATE_TYPE then coalesce(sd.defaultFmts.dateFmt, ctx.defaultFmts.dateFmt, DEFAULT_DATE_FMT)
           when dbms_sql.TIMESTAMP_TYPE then coalesce(sd.defaultFmts.timestampFmt, ctx.defaultFmts.timestampFmt, DEFAULT_TIMESTAMP_FMT)
           when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then coalesce(sd.defaultFmts.timestampFmt, ctx.defaultFmts.timestampFmt, DEFAULT_TIMESTAMP_FMT)
           end;    
  end;
  
  function makeCellRef (
    colRef  in varchar2
  , rowRef  in pls_integer
  )
  return cell_ref_t
  is
    cellRef  cell_ref_t;
  begin
    if colRef is null and rowRef is not null then
      error('column reference argument cannot be null');
    elsif colRef is not null and rowRef is null then
      error('row reference argument cannot be null');
    end if;
    cellRef.c := colRef;
    cellRef.cn := base26decode(cellRef.c);
    if cellRef.cn not between 1 and MAX_COLUMN_NUMBER then
      error(RANGE_INVALID_COL, cellRef.c);
    end if;
    if rowRef not between 1 and MAX_ROW_NUMBER then
      error(RANGE_INVALID_ROW, rowRef);
    end if;
    cellRef.r := rowRef;
    cellRef.value := cellRef.c || to_char(cellRef.r);
    return cellRef;
  end;

  function makeCellRef (
    p_colIdx  in pls_integer
  , p_rowIdx  in pls_integer
  )
  return varchar2
  is
  begin
    return makeCellRef(base26encode(p_colIdx), p_rowIdx).value;
  end;

  function translateCellRef (
    cellRef   in cell_ref_t
  , colShift  in pls_integer default 0
  , rowShift  in pls_integer default 0
  )
  return cell_ref_t
  is
    newCellRef  cell_ref_t := cellRef;
  begin
    newCellRef.cn := cellRef.cn + nvl(colShift, 0);
    if newCellRef.cn not between 1 and MAX_COLUMN_NUMBER then
      error('Column index out of range: %d', newCellRef.cn);
    else
      newCellRef.c := base26encode(newCellRef.cn);
    end if;
    newCellRef.r := cellRef.r + nvl(rowShift, 0);
    if newCellRef.r not between 1 and MAX_ROW_NUMBER then
      error('Row index out of range: %d', newCellRef.r);
    end if;
    newCellRef.value := newCellRef.c || to_char(newCellRef.r);
    return newCellRef;
  end;
  
  function makeRangeImpl (
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
    if range.start_ref.value is null then
      error(RANGE_EMPTY_REF);
    end if;
    range.end_ref := makeCellRef(endCol, endRow);
    range.expr := range.start_ref.value || case when range.end_ref.value is not null then ':'||range.end_ref.value end;
    range.is_null := false;
    return range;
  end;
  
  function makeRangeImpl (
    cellSpan  in cellSpan_t 
  )
  return range_t
  is
  begin
    return makeRangeImpl( 
             base26encode(cellSpan.anchorRef.colOffset)
           , cellSpan.anchorRef.rowOffset
           , base26encode(cellSpan.anchorRef.colOffset + cellSpan.colSpan - 1)
           , cellSpan.anchorRef.rowOffset + cellSpan.rowSpan - 1
           );
  end;
  
  function makeCellRange (
    p_startRowIdx  in pls_integer
  , p_startColIdx  in pls_integer
  , p_endRowIdx    in pls_integer default null
  , p_endColIdx    in pls_integer default null
  )
  return ExcelTypes.ST_Ref
  is
  begin
    return makeRangeImpl(
             startCol  => base26encode(p_startColIdx)
           , startRow  => p_startRowIdx
           , endCol    => base26encode(p_endColIdx)
           , endRow    => p_endRowIdx
           ).expr;
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
      cellRef.value := expr;
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
    
    range.expr := expr;
    range.is_null := false;
    
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
  
  function blobToXml (
    input  in blob
  ) 
  return XMLType 
  is
    -- Byte order marks for UTF-8, UTF-16LE and UTF-16BE
    C_BOM_UTF8     constant raw(3) := hextoraw('EFBBBF');
    C_BOM_UTF16LE  constant raw(2) := hextoraw('FFFE');
    C_BOM_UTF16BE  constant raw(2) := hextoraw('FEFF');
    
    segmentStart       integer;
    segmentStop        integer := 0;
    segmentSize        pls_integer;
    xmlBlob            blob := input;
    xmlOutput          xmltype;
    xmlProlog          varchar2(32767);
    xmlEncoding        varchar2(30) := 'UTF-8';
    oracleCharsetName  varchar2(30);
    bom                raw(3);
    
  begin
    
    bom := dbms_lob.substr(xmlBlob, 3);
    
    if bom = C_BOM_UTF8 then   
      xmlEncoding := 'UTF-8';
    else
      
      case utl_raw.substr(bom, 1, 2)
      when C_BOM_UTF16LE then
        xmlEncoding := 'UTF-16LE';
      
      when C_BOM_UTF16BE then
        xmlEncoding := 'UTF-16BE';
        
      else
      
        -- search for an XML prolog
        segmentStart := dbms_lob.instr(xmlBlob, utl_raw.cast_to_raw('<?xml'));
        
        if segmentStart != 0 then
          
          segmentStop := dbms_lob.instr(xmlBlob, utl_raw.cast_to_raw('?>'), segmentStart) + 1;
          segmentSize := segmentStop - segmentStart + 1;
          xmlProlog := utl_raw.cast_to_varchar2(dbms_lob.substr(xmlBlob, segmentSize, segmentStart));
          
          -- search for a well-formed XML declaration in the prolog 
          -- as per https://www.w3.org/TR/REC-xml/#NT-XMLDecl
          xmlEncoding := regexp_substr(
                           xmlProlog
                         , '<\?xml'
                         || '\s+version\s*=\s*(''1\.[0-9]+''|"1\.[0-9]+")'     -- version
                         || '(\s+encoding\s*=\s*(''.*?''|".*?"))?'             -- encoding
                         || '(\s+standalone\s*=\s*(''(yes|no)''|"(yes|no)"))?' -- standalone  
                         || '\s*\?>'
                         , 1
                         , 1
                         , null
                         , 3
                         );
          
          -- trim apostrophes or quotes
          xmlEncoding := substr(xmlEncoding, 2, length(xmlEncoding)-2);
          
        end if;
        
        -- search for a DOCTYPE declaration
        segmentStart := dbms_lob.instr(xmlBlob, utl_raw.cast_to_raw('<!DOCTYPE'), segmentStop + 1);
        if segmentStart != 0 then
          segmentStop := dbms_lob.instr(xmlBlob, utl_raw.cast_to_raw('>'), segmentStart);
          segmentSize := segmentStop - segmentStart + 1;
          -- copy and overwrite with blanks
          dbms_lob.createtemporary(xmlBlob, true);
          dbms_lob.copy(xmlBlob, input, dbms_lob.lobmaxsize);
          dbms_lob.write(xmlBlob, segmentSize, segmentStart, utl_raw.copies('20', segmentSize));
        end if;
        
      end case;    
        
    end if;
    
    oracleCharsetName := utl_i18n.map_charset(xmlEncoding, flag => utl_i18n.IANA_TO_ORACLE);
    xmlOutput := xmltype(xmlBlob, nls_charset_id(oracleCharsetName));
    if dbms_lob.istemporary(xmlBlob) = 1 then
      dbms_lob.freetemporary(xmlBlob);
    end if;
    
    return xmlOutput;

  end;

  procedure string_write (
    buf  in out nocopy varchar2
  , str  in varchar2
  )
  is
  begin
    buf := buf || str;
  end;
  
  function makeRgbColor (
    r  in uint8
  , g  in uint8
  , b  in uint8
  , a  in number default null
  )
  return varchar2
  is
  begin
    return '#' || ExcelTypes.makeRgbColor(r,g,b,a);
  end;

  function getNextNumFmtId return pls_integer is
  begin
    currentCtx.workbook.styles.numFmtIdSeq := nvl(currentCtx.workbook.styles.numFmtIdSeq, 163) + 1;
    return currentCtx.workbook.styles.numFmtIdSeq;
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
      --numFmtId := nvl(styles.numFmts.last, 163) + 1;
      numFmtId := getNextNumFmtId;
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
  begin
    return ExcelTypes.makeBorderPr(p_style, p_color);
  end;
  
  function makeBorder (
    p_left    in CT_BorderPr default makeBorderPr()
  , p_right   in CT_BorderPr default makeBorderPr()
  , p_top     in CT_BorderPr default makeBorderPr()
  , p_bottom  in CT_BorderPr default makeBorderPr()
  )
  return CT_Border
  is
  begin
    return ExcelTypes.makeBorder(p_left, p_right, p_top, p_bottom);
  end;
  
  function makeBorder (
    p_style  in varchar2
  , p_color  in varchar2 default null
  )
  return CT_Border
  is
  begin
    return ExcelTypes.makeBorder(p_style, p_color);
  end;
  
  function putBorder (
    styles  in out nocopy CT_Stylesheet
  , border  in CT_Border
  )
  return pls_integer
  is
    borderId  pls_integer := 0;
  begin
    if border.content is not null then
      if styles.borderMap.exists(border.content) then
        borderId := styles.borderMap(border.content);
      else
        borderId := nvl(styles.borders.last, -1) + 1;
        styles.borders(borderId) := border;
        styles.borderMap(border.content) := borderId;
      end if;
    end if;
    return borderId;
  end;

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
  return CT_Font
  is
  begin
    return ExcelTypes.makeFont(p_name, p_sz, p_b, p_i, p_color, p_u, p_vertAlign, p_strike);
  end;
  
  function putFont (
    styles  in out nocopy CT_Stylesheet
  , font    CT_Font
  )
  return pls_integer
  is
    fontId  pls_integer := 0;
  begin
    if font.content is not null then
      if styles.fontMap.exists(font.content) then
        fontId := styles.fontMap(font.content);
      else
        fontId := nvl(styles.fonts.last, -1) + 1;
        styles.fonts(fontId) := font;
        styles.fontMap(font.content) := fontId;
      end if;
    end if;
    return fontId;
  end;

  function makePatternFill (
    p_patternType  in varchar2
  , p_fgColor      in varchar2 default null
  , p_bgColor      in varchar2 default null
  )
  return CT_Fill
  is
  begin
    return ExcelTypes.makePatternFill(p_patternType, p_fgColor, p_bgColor);
  end;

  function makeGradientStop (
    p_position  in number
  , p_color     in varchar2
  )
  return CT_GradientStop
  is
  begin
    return ExcelTypes.makeGradientStop(p_position, p_color);
  end;
  
  function makeGradientFill (
    p_degree  in number default null
  , p_stops   in CT_GradientStopList default null
  )
  return CT_Fill
  is
  begin
    return ExcelTypes.makeGradientFill(p_degree, p_stops);
  end;

  procedure addGradientStop (
    p_fill      in out nocopy CT_Fill
  , p_position  in number
  , p_color     in varchar2
  )
  is
  begin
    ExcelTypes.addGradientStop(p_fill, p_position, p_color);
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

  function makeAlignment (
    p_horizontal    in varchar2 default null
  , p_vertical      in varchar2 default null
  , p_wrapText      in boolean default false
  , p_textRotation  in number default null
  , p_verticalText  in boolean default false
  , p_indent        in number default null
  )
  return CT_CellAlignment
  is
  begin
    return ExcelTypes.makeAlignment(p_horizontal, p_vertical, p_wrapText, p_textRotation, p_verticalText, p_indent);
  end;

  procedure setCellXfContent (
    xf  in out nocopy CT_Xf
  )
  is
  begin
    
    xf.content := null;
    
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
      --xf.content := null;
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
    
    setCellXfContent(xf);
    
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

  function getCellXf (
    ctx   in context_t
  , xfId  in pls_integer
  )
  return CT_Xf
  is
  begin
    return ctx.workbook.styles.cellXfs(xfId);
  end;

  function getCellStyleXf (
    ctx   in context_t
  , xfId  in pls_integer
  )
  return CT_Xf
  is
  begin
    return ctx.workbook.styles.cellStyleXfs(xfId);
  end;

  function getCellFont (
    ctx   in context_t
  , xfId  in pls_integer    
  )
  return CT_Font
  is
    xf  CT_Xf := getCellXf(ctx, xfId);
  begin
    return ctx.workbook.styles.fonts(xf.fontId);
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
    xf  CT_Xf;
  begin
    loadContext(p_ctxId);
    xf := makeCellXf(currentCtx.workbook.styles, 0, p_numFmtCode, p_font, p_fill, p_border, p_alignment);
    return putCellXf(currentCtx.workbook.styles, xf);
  end;

  function makeCellStyleCss (
    p_ctxId  in ctxHandle
  , p_css    in varchar2
  )
  return cellStyleHandle
  is
    style  ExcelTypes.CT_Style := ExcelTypes.getStyleFromCss(p_css);
    xf     CT_Xf;
  begin
    loadContext(p_ctxId);
    xf := makeCellXf(currentCtx.workbook.styles, 0, style.numberFormat, style.font, style.fill, style.border, style.alignment);
    return putCellXf(currentCtx.workbook.styles, xf);
  end;

  function putDxf (
    styles  in out nocopy CT_Stylesheet
  , style   in ExcelTypes.CT_Style
  )
  return pls_integer
  is
    dxfId     pls_integer;
    dxf       CT_Dxf;
  begin
    dxf.style := style;
    
    -- font
    dxf.content := dxf.style.font.content;
    
    -- numFmt
    if dxf.style.numberFormat is not null then
      -- numFmtMap is shared between xf and dxf
      if styles.numFmtMap.exists(dxf.style.numberFormat) then
        dxf.style.numFmtId := styles.numFmtMap(dxf.style.numberFormat);
      else 
        dxf.style.numFmtId := getNextNumFmtId;
        styles.numFmtMap(dxf.style.numberFormat) := dxf.style.numFmtId;
      end if;
      dxf.content := dxf.content || 
                     '<numFmt numFmtId="'||to_char(dxf.style.numFmtId)||'" formatCode="'||dbms_xmlgen.convert(dxf.style.numberFormat)||'"/>' ;
    end if;
    
    -- fill
    -- ugly workaround: 
    -- in a dxf context, background color for solid pattern is stored in the bgColor element, contrary to regular xf cell style which uses fgColor
    if dxf.style.fill.fillType = ExcelTypes.FT_PATTERN and dxf.style.fill.patternFill.patternType = 'solid' then
      dxf.style.fill.patternFill.bgColor := null;
      ExcelTypes.swapPatternFillColors(dxf.style.fill);
    end if;
    dxf.content := dxf.content || dxf.style.fill.content;
    
    -- alignment
    dxf.content := dxf.content || dxf.style.alignment.content;
    
    -- border
    dxf.content := dxf.content || dxf.style.border.content;
    
    if styles.dxfMap.exists(dxf.content) then
      dxfId := styles.dxfMap(dxf.content);
    else
      dxfId := nvl(styles.dxfs.last, -1) + 1;
      styles.dxfMap(dxf.content) := dxfId;
      styles.dxfs(dxfId) := dxf;
    end if;
    
    return dxfId;
  end;
  
  function makeCondFmtStyle (
    p_ctxId       in ctxHandle
  , p_numFmtCode  in varchar2 default null
  , p_font        in CT_Font default null
  , p_fill        in CT_Fill default null
  , p_border      in CT_Border default null
  , p_alignment   in CT_CellAlignment default null
  )
  return cfmtStyleHandle
  is
    style  ExcelTypes.CT_Style;
  begin
    loadContext(p_ctxId);
    
    style.numberFormat := p_numFmtCode;
    style.font := p_font;
    style.fill := p_fill;
    style.border := p_border;
    style.alignment := p_alignment;
        
    return putDxf(currentCtx.workbook.styles, style);
  end;

  function makeCondFmtStyleCss (
    p_ctxId  in ctxHandle
  , p_css    in varchar2
  )
  return cfmtStyleHandle
  is
  begin
    loadContext(p_ctxId);
    return putDxf(currentCtx.workbook.styles, ExcelTypes.getStyleFromCss(p_css));
  end;

  function mergeCellFormat (
    ctx     in out nocopy context_t
  , style   in cellStyleHandle
  , format  in varchar2
  , force   in boolean default false
  )
  return cellStyleHandle
  is
    xf    CT_Xf;
    xfId  pls_integer := style;
  begin
    if xfId is not null then
      -- get style definition record
      xf := getCellXf(ctx, xfId);
      -- set format property
      if format is not null and ( xf.numFmtId = 0 or force ) then
        xf.numFmtId := putNumfmt(ctx.workbook.styles, format);
        setCellXfContent(xf); -- update content
        xfId := putCellXf(ctx.workbook.styles, xf);
      end if;      
    else
      xf := makeCellXf(ctx.workbook.styles, 0, format);
      xfId := putCellXf(ctx.workbook.styles, xf);
    end if;
    return xfId;
  end;

  procedure mergeCellStyleImpl (
    ctx         in out nocopy context_t
  , masterXf    in CT_Xf
  , xf          in out nocopy CT_Xf
  )
  is
    style        ExcelTypes.CT_Style;
  begin 
  
    if xf.xfId = 0 then
      xf.xfId := masterXf.xfId;
    end if;
      
    -- number format
    if xf.numFmtId = 0 then
      xf.numFmtId := masterXf.numFmtId;
    end if;
      
    -- font
    if xf.fontId != 0 then
      style.font := ExcelTypes.mergeFonts( ctx.workbook.styles.fonts(masterXf.fontId)
                                         , ctx.workbook.styles.fonts(xf.fontId) );
      xf.fontId := putFont(ctx.workbook.styles, style.font); 
    else
      xf.fontId := masterXf.fontId;
    end if;
      
    -- fill
    if xf.fillId != 0 then
      style.fill := ctx.workbook.styles.fills(xf.fillId);
      if style.fill.fillType = ExcelTypes.FT_PATTERN then
        style.fill := ExcelTypes.mergePatternFills( ctx.workbook.styles.fills(masterXf.fillId)
                                                  , style.fill );
        xf.fillId := putFill(ctx.workbook.styles, style.fill);
      -- else, should be a gradientFill
      end if;
    else
      xf.fillId := masterXf.fillId;
    end if;
      
    -- border
    if xf.borderId != 0 then
      style.border := ExcelTypes.mergeBorders( ctx.workbook.styles.borders(masterXf.borderId)
                                             , ctx.workbook.styles.borders(xf.borderId) );
      xf.borderId := putBorder(ctx.workbook.styles, style.border);
    else
      xf.borderId := masterXf.borderId;
    end if;
      
    -- alignment
    if xf.alignment.content is not null then
      xf.alignment := ExcelTypes.mergeAlignments(masterXf.alignment, xf.alignment);
    else
      xf.alignment := masterXf.alignment;
    end if;
    
  end;

  function mergeCellStyle (
    ctx         in out nocopy context_t
  , masterXfId  in cellStyleHandle
  , xfId        in cellStyleHandle
  )
  return cellStyleHandle
  is
    xf        CT_Xf;
  begin
    if xfId != 0 then
      if masterXfId != 0 then
        xf := getCellXf(ctx, xfId);
        mergeCellStyleImpl(ctx, getCellXf(ctx, masterXfId), xf);
        setCellXfContent(xf);
        return putCellXf(ctx.workbook.styles, xf);
      else
        return xfId;
      end if;
    else
      return masterXfId;
    end if;
  end;
  
  function mergeLinkFont (
    ctx       in out nocopy context_t
  , linkXfId  in cellStyleHandle
  , xfId      in cellStyleHandle
  )
  return cellStyleHandle
  is
    xf  CT_Xf;
  begin
    if xfId != 0 then
      if linkXfId != 0 then
        xf := getCellXf(ctx, xfId);
        xf.xfId := linkXfId;
        xf.fontId := getCellStyleXf(ctx, linkXfId).fontId;
        setCellXfContent(xf);
        return putCellXf(ctx.workbook.styles, xf);
      else
        return xfId;
      end if;
    else
      return linkXfId;
    end if;
  end;
  
  function setRangeBorders (
    xfId      in pls_integer
  , cellSpan  in cellSpan_t
  , rowIdx    in pls_integer
  , colIdx    in pls_integer
  )
  return pls_integer
  is
    xf      CT_Xf := getCellXf(currentCtx, xfId);
    border  CT_Border:= currentCtx.workbook.styles.borders(xf.borderId);
  begin
    if xf.borderId != 0 then
      -- apply outside border
      border := ExcelTypes.applyBorderSide(
                  border => border
                , top    => rowIdx = cellSpan.anchorRef.rowOffset
                , right  => colIdx = cellSpan.anchorRef.colOffset + cellSpan.colSpan - 1
                , bottom => rowIdx = cellSpan.anchorRef.rowOffset + cellSpan.rowSpan - 1
                , left   => colIdx = cellSpan.anchorRef.colOffset
                );
      xf.borderId := putBorder(currentCtx.workbook.styles, border);
      setCellXfContent(xf);
      return putCellXf(currentCtx.workbook.styles, xf);
    else
      return xfId;
    end if;
  end;

  function putRichValueStructure (
    rvStructures  in out nocopy ExcelTypes.CT_RichValueStructures
  , struct        in ExcelTypes.CT_RichValueStructure
  )
  return pls_integer 
  is
    structId  pls_integer;
  begin
    if rvStructures.structMap.exists(struct.t) then
      structId := rvStructures.structMap(struct.t);
    else
      rvStructures.structs.extend;
      structId := rvStructures.structs.last;
      rvStructures.structs(structId) := struct;
      structId := structId - 1; -- transform to 0-based index
      rvStructures.structMap(struct.t) := structId;
    end if;
    return structId;
  end;
  
  function putLocalImageRvStruct (
    rvStructures  in out nocopy ExcelTypes.CT_RichValueStructures
  )
  return pls_integer
  is
    struct  ExcelTypes.CT_RichValueStructure;
  begin
    struct.t := '_localImage';
    struct.keys := ExcelTypes.CT_KeyList(null,null);
    struct.keys(1).n := '_rvRel:LocalImageIdentifier';
    struct.keys(1).t := 'i';
    struct.keys(2).n := 'CalcOrigin';
    struct.keys(2).t := 'i';
    return putRichValueStructure(rvStructures, struct);
  end;

  function colPxToCharWidth (
    p_px  in pls_integer
  )
  return number
  is
  begin
    return case when p_px < 12 then p_px/12
                else trunc((p_px - 5)/7 * 100 + .5)/100
           end;
  end;

  function rowPxToPt (
    p_px  in pls_integer
  )
  return number
  is
  begin
    return p_px * 3 / 4;
  end;

  function getColumnWidth (
    displayWidth in number 
  )
  return binary_double
  is
  begin
    return to_binary_double(
             trunc(
               round(
                 case when displayWidth < 1 then
                   -- for display width less than 1 char unit in the default Normal font,
                   -- the internal computed width is directly proportional to the width in pixel of a 1-char column (7+5)
                   displayWidth * 12
                 else
                   displayWidth * 7 + 5
                 end
               ) -- rounding to get an integer number of pixels
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
    meta         in out nocopy sql_metadata_t
  , ctxColumnId  in pls_integer
  )
  is
    
    col           column_t := meta.columnList(ctxColumnId);
    tokenId       pls_integer := 0;
    token         varchar2(128);
    tokenMap      column_Map_t;
    refColumnId   pls_integer;
    refColumn     column_t;

    function next_token return varchar2 is
    begin
      tokenId := tokenId + 1;
      return regexp_substr(col.fmla.expr, '\$\{([^}]+)\}', 1, tokenId, null, 1);
    end;
    
  begin
    
    -- parse tokens
    token := next_token;
    while token is not null loop
      if meta.columnMap.exists(token) then
        tokenMap(token) := meta.columnMap(token);
      else
        error('Unknown column name in hyperlink token: %s', token);
      end if;
      token := next_token;  
    end loop;
    
    col.linkTokens := intList_t();
    
    token := tokenMap.first;
    while token is not null loop  
      -- get column instance referenced by this token
      refColumnId := tokenMap(token);
      refColumn := meta.columnList(refColumnId);
      if not(refColumn.excluded or refColumn.id = col.id) then
        -- replace token occurrence(s) with the cell reference in R1C1 style
        col.fmla.expr := replace(col.fmla.expr, '${'||token||'}', '"&RC['||to_char(refColumn.id - col.id)||']&"');
      else
        -- replace with a numeric token, a literal substitution will be performed later on when the actual column value is known
        col.fmla.expr := replace(col.fmla.expr, '${'||token||'}', '${'||to_char(refColumnId)||'}');
        col.linkTokens.extend;
        col.linkTokens(col.linkTokens.last) := refColumnId;
      end if;
      token := tokenMap.next(token);
    end loop;
    
    -- clean up leading and trailing empty strings
    col.fmla.expr := regexp_replace(col.fmla.expr, '""&|&""');
    debug(col.fmla.expr);
    
    -- having substitutable tokens make this formula unshareable
    if col.linkTokens.count != 0 then
      col.fmla.shared := false;
      col.fmla.sharedIdx := null;
    end if;
    
    meta.columnList(ctxColumnId) := col;
    
  end;
  
  procedure prepareHyperlinks (
    sd  in out nocopy sheet_definition_t
  )
  is
    columnId  pls_integer;
  begin
    for i in 1 .. sd.tableList.count loop     
      columnId := sd.tableList(i).columnLinkMap.first;
      while columnId is not null loop       
        --prepareHyperlink(sd.tableList(i).columnLinkMap(columnId), columnId, sd.tableList(i).sqlMetadata.columnList);     
        columnId := sd.tableList(i).columnLinkMap.next(columnId);
      end loop;
    end loop;
  end;

  procedure setLinkTokenValues (
    expr     in out nocopy varchar2
  , tokens   in intList_t
  , dataMap  in data_map_t
  )
  is
  begin
    for i in 1 .. tokens.count loop
      expr := replace(expr, '${'||to_char(tokens(i))||'}', dataMap(tokens(i)).varchar2_value);
    end loop;
    debug(expr);
  end;

  function newStylesheet
  return CT_Stylesheet
  is
    styles  CT_Stylesheet;
    dummy   pls_integer;
    xfId    pls_integer;
  begin
    dummy := putFont(styles, makeFont(ExcelTypes.DEFAULT_FONT_FAMILY, ExcelTypes.DEFAULT_FONT_SIZE));
    dummy := putFill(styles, makePatternFill('none'));
    dummy := putFill(styles, makePatternFill('gray125'));
    dummy := putBorder(styles, makeBorder('none'));
    
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
    ctx       in out nocopy context_t
  , str       in varchar2
  , richText  in boolean default false
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
      ctx.string_list(idx).value := str;
      ctx.string_list(idx).isRichText := richText;
    else
      idx := ctx.string_map(str);
    end if;
    return idx;
  end;

  function put_rt (
    ctx  in out nocopy context_t
  , rt   in ExcelTypes.CT_RichText
  ) 
  return pls_integer
  is
    idx  pls_integer := put_string(ctx, rt.content, true);
  begin
    ctx.rt_cache(idx) := rt;
    return idx;
  end;

  function getCursorNumber (
    p_query in clob --varchar2 
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
    p_cursor_number   in integer
  , p_excludeSet      in valueSet_t
  , p_offset          in pls_integer
  , p_virtualColumns  in virtualColumnList_t
  )
  return column_list_t
  is
    dbColumnList    dbms_sql.desc_tab3;
    columnCount     integer;
    data            data_t;
    columnList      column_list_t := column_list_t();
    COLUMN_DEFAULT  column_t;
    col             column_t;
    columnSeq       pls_integer := 0;
    cols            dbLinkedList_t;
    vc              virtualColumn_t;
    nodeId          pls_integer;
    
    function node (data in column_t) return pls_integer is
      node  dbLinkedNode_t;
    begin
      node.id := cols.heap.count + 1;
      
      if node.data.excluded then
        node.id := node.id * -1;
      end if;
      
      node.data := data;
      cols.heap(node.id) := node;
      return node.id;
    end;
    
  begin
    dbms_sql.describe_columns3(p_cursor_number, columnCount, dbColumnList);
    
    for i in 1 .. columnCount loop
      
      col := COLUMN_DEFAULT;
    
      case dbColumnList(i).col_type
      when dbms_sql.VARCHAR2_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.varchar2_value, dbColumnList(i).col_max_len);
        col.supertype := ST_STRING;
      when dbms_sql.CHAR_TYPE then
        dbms_sql.define_column_char(p_cursor_number, i, data.char_value, dbColumnList(i).col_max_len);
        col.supertype := ST_STRING;
      when dbms_sql.NUMBER_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.number_value);
        col.supertype := ST_NUMBER;
      when dbms_sql.DATE_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.date_value);
        col.supertype := ST_DATETIME;
      when dbms_sql.TIMESTAMP_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.ts_value);
        col.supertype := ST_DATETIME;
      when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.tstz_value);
        col.supertype := ST_DATETIME;
      when dbms_sql.CLOB_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.clob_value);
        col.supertype := ST_LOB;
      when dbms_sql.USER_DEFINED_TYPE then
        if dbColumnList(i).col_type_name = 'ANYDATA' then
          dbms_sql.define_column(p_cursor_number, i, data.anydata_value);
          col.supertype := ST_VARIANT;
        else
          error('Unsupported data type: %d [%s], for column "%s"', dbColumnList(i).col_type, dbColumnList(i).col_type_name, dbColumnList(i).col_name);
        end if;
      when dbms_sql.BLOB_TYPE then
        dbms_sql.define_column(p_cursor_number, i, data.blob_value);
        col.supertype := ST_IMAGE;
      else
        error('Unsupported data type: %d, for column "%s"', dbColumnList(i).col_type, dbColumnList(i).col_name);
      end case;
      
      col.dbId := i;
      col.name := dbColumnList(i).col_name;
      col.type := dbColumnList(i).col_type;
      col.scale := dbColumnList(i).col_scale;
      col.excluded := ( p_excludeSet.integers.exists(i) or p_excludeSet.strings.exists(col.name) );
      
      insertLast(cols, node(col));
      
    end loop;
    
    for i in 1 .. p_virtualColumns.count loop
    
      vc := p_virtualColumns(i);
      
      if vc.pos is null then
        insertLast(cols, node(vc.col));
      else
        if vc.after then
          insertAfter(cols, vc.pos, node(vc.col));
        else
          insertBefore(cols, vc.pos, node(vc.col));
        end if;
      end if;
    
    end loop;
    
    nodeId := cols.first;
    while nodeId is not null loop
      
      col := cols.heap(nodeId).data;
      
      if not col.excluded then
        columnSeq := columnSeq + 1;
        col.id := columnSeq;
        col.colNum := p_offset - 1 + columnSeq;
        col.colRef := base26encode(col.colNum);
      end if;
      
      columnList.extend;
      columnList(columnList.last) := col;
      
      nodeId := cols.heap(nodeId).next;
      
    end loop;
    
    return columnList;
  end;

  procedure prepareCursor (
    meta       in out nocopy sql_metadata_t
  , colOffset  in pls_integer
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
    
    meta.columnList := getColumnList(meta.cursorNumber, meta.excludeSet, colOffset, meta.virtualColumns);
    
    for i in 1 .. meta.columnList.count loop
      meta.columnMap(meta.columnList(i).name) := i;
      if not meta.columnList(i).excluded then
        meta.visibleColumnSet(i) := meta.columnList(i).colRef;
      end if;
    end loop;
    
  end;
  
  procedure closeCursor (meta in out nocopy sql_metadata_t)
  is
  begin
    debug('close cursor');
    dbms_sql.close_cursor(meta.cursorNumber);
  end;

  procedure prepareNumberValue (data in out nocopy data_t, v in number)
  is
  begin
    data.st := ST_NUMBER;
    data.db_type := dbms_sql.NUMBER_TYPE;
    data.number_value := v;
    data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
  end;

  procedure prepareStringValue (data in out nocopy data_t, v in varchar2)
  is
  begin
    data.st := ST_STRING;
    data.db_type := dbms_sql.VARCHAR2_TYPE;
    data.varchar2_value := stripXmlControlChars(v);
  end;

  procedure prepareDateValue (data in out nocopy data_t, v in date)
  is
  begin
    data.st := ST_DATETIME;
    data.db_type := dbms_sql.DATE_TYPE;
    data.date_value := v;
    data.number_value := toOADate(dt => data.date_value);
    data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
  end;

  procedure prepareTimestampValue (data in out nocopy data_t, v in timestamp_unconstrained)
  is
  begin
    data.st := ST_DATETIME;
    data.db_type := dbms_sql.TIMESTAMP_TYPE;
    data.ts_value := timestampRound(v, 3);
    data.number_value := toOADate(ts => data.ts_value);
    data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
  end;

  procedure prepareTimestampTzValue (data in out nocopy data_t, v in timestamp_tz_unconstrained)
  is
  begin
    data.st := ST_DATETIME;
    data.db_type := dbms_sql.TIMESTAMP_WITH_TZ_TYPE;
    data.tstz_value := timestampRound(v, 3);
    data.number_value := toOADate(ts => data.tstz_value);
    data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
  end;

  procedure prepareData (
    data  in out nocopy data_t
  , v     in anydata
  ) 
  is
  begin
    case v.GetTypeName()
    when 'SYS.NUMBER' then
      prepareNumberValue(data, v.AccessNumber());
    when 'SYS.VARCHAR2' then
      prepareStringValue(data, v.AccessVarchar2());
    when 'SYS.CHAR' then
      prepareStringValue(data, rtrim(v.AccessChar()));
    when 'SYS.DATE' then
      prepareDateValue(data, v.AccessDate());
    when 'SYS.TIMESTAMP' then
      prepareTimestampValue(data, v.AccessTimestamp());
    when 'SYS.TIMESTAMP_WITH_TIMEZONE' then
      prepareTimestampTzValue(data, v.AccessTimestampTZ());
    when 'SYS.CLOB' then
      data.db_type := dbms_sql.CLOB_TYPE;
      data.clob_value := v.AccessClob();
      data.st := ST_LOB;
    when 'SYS.BLOB' then
      data.db_type := dbms_sql.BLOB_TYPE;
      data.blob_value := v.AccessBlob();
      data.st := ST_IMAGE;
    else
      error('Unsupported data type: ''%s''', v.GetTypeName());
    end case;
  end;

  function getSqlData (
    sqlMeta  sql_metadata_t
  )
  return data_map_t
  is
    dbId     pls_integer;
    data     data_t;
    dataMap  data_map_t;
  begin

    for i in 1 .. sqlMeta.columnList.count loop
                    
      data := null;
      data.st := sqlMeta.columnList(i).supertype;
      data.db_type := sqlMeta.columnList(i).type;
      dbId := sqlMeta.columnList(i).dbId;

      if not sqlMeta.columnList(i).virtual then

        case data.db_type
        when dbms_sql.VARCHAR2_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.varchar2_value);
          data.varchar2_value := stripXmlControlChars(data.varchar2_value);
              
        when dbms_sql.CHAR_TYPE then
          dbms_sql.column_value_char(sqlMeta.cursorNumber, dbId, data.char_value);
          data.varchar2_value := stripXmlControlChars(rtrim(data.char_value));
              
        when dbms_sql.NUMBER_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.number_value);
          if sqlMeta.columnList(i).scale between -84 and 0 then
            data.varchar2_value := to_char(data.number_value);
          else
            data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
          end if;
              
        when dbms_sql.DATE_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.date_value);
          data.number_value := toOADate(dt => data.date_value);
          data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
              
        when dbms_sql.TIMESTAMP_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.ts_value);
          data.ts_value := timestampRound(data.ts_value, 3);
          data.number_value := toOADate(ts => data.ts_value);
          data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
              
        when dbms_sql.TIMESTAMP_WITH_TZ_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.tstz_value);
          data.tstz_value := timestampRound(data.tstz_value, 3);
          data.number_value := toOADate(ts => data.tstz_value);
          data.varchar2_value := to_char(data.number_value, 'TM9', NLS_PARAM_STRING);
              
        when dbms_sql.CLOB_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.clob_value);
          
        when dbms_sql.USER_DEFINED_TYPE then -- should be ANYDATA
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.anydata_value);
          prepareData(data, data.anydata_value);
        
        when dbms_sql.BLOB_TYPE then
          dbms_sql.column_value(sqlMeta.cursorNumber, dbId, data.blob_value);
        
        end case;
      
      else
        
        null;
      
      end if;
          
      dataMap(i) := data;
          
    end loop;
    
    return dataMap;
    
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

  function getTableForest (tableList in tableList_t) return tableForest_t
  is
    f              tableForest_t;
    anchorTableId  pls_integer;
    
    procedure push (list in out nocopy intList_t, v in pls_integer) is
    begin
      list.extend;
      list(list.last) := v;
    end;
    
  begin
    f.roots := intList_t();
  
    for i in 1 .. tableList.count loop
      f.t(i).nodeId := i;
      f.t(i).children := intList_t();
      anchorTableId := tableList(i).anchorRef.tableId;
      if anchorTableId is not null then
        push(f.t(anchorTableId).children, i);
      else
        push(f.roots, i);
      end if;
    end loop;
    
    return f;
  end;
  
  function putDefaultExtension (
    ctx          in out nocopy context_t
  , contentType  in varchar2
  )
  return varchar2
  is
    ext  varchar2(16);
  begin
    if ctx.extensions.exists(contentType) then
      ext := ctx.extensions(contentType);
    else
      ext := case contentType
             when MT_PNG then 'png'
             when MT_JPEG then 'jpeg'
             when MT_GIF then 'gif'
             when MT_SVG then 'svg'
             --TODO
             end;
      ctx.extensions(contentType) := ext;
    end if;
    return ext;
  end;

  function new_part (
    name  in varchar2
  , contentType  in varchar2
  )
  return part_t
  is
    p  part_t;
  begin
    p.name := name;
    p.contentType := contentType;
    p.rels := CT_Relationships();
    return p;
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
    part  part_t := new_part(name, contentType);
  begin
    part.content := content;
    addPart(ctx, part);
  end;

  procedure addPart (
    ctx          in out nocopy context_t
  , name         in varchar2
  , contentType  in varchar2
  , contentBin   in blob
  )
  is
    part  part_t := new_part(name, contentType);
  begin
    part.contentBin := contentBin;
    part.isBinary := true;
    addPart(ctx, part);
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

  function addTableLayout (
    ctx                in out nocopy context_t
  , tableRange         in range_t
  , showHeader         in boolean
  , tableAutoFilter    in boolean
  , tableStyleName     in varchar2
  , columnMap          in table_column_map_t
  , tableName          in varchar2 default null
  , showFirstColumn    in boolean
  , showLastColumn     in boolean
  , showRowStripes     in boolean
  , showColumnStripes  in boolean
  )
  return pls_integer
  is
    tab      CT_Table;
  begin
    tab.id := nvl(ctx.workbook.tables.last, 0) + 1;
    if tableName is null then  
      loop
        ctx.tableNameSeq := ctx.tableNameSeq + 1;
        tab.name := 'Table'||to_char(ctx.tableNameSeq);
        exit when not ctx.nameMap.exists(upper(tab.name));
      end loop;
      ctx.nameMap(upper(tab.name)) := null;
    else
      tab.name := tableName;
    end if;
    
    tab.ref := tableRange;
    tab.showHeader := nvl(showHeader, false);
    tab.autoFilter := nvl(tableAutoFilter, false);
    tab.styleName := tableStyleName;
    tab.showFirstColumn := nvl(showFirstColumn, false);
    tab.showLastColumn := nvl(showLastColumn, false);
    tab.showRowStripes := nvl(showRowStripes, true);
    tab.showColumnStripes := nvl(showColumnStripes, false);
    
    tab.partName := 'xl/tables/table'||to_char(tab.id)
                                     ||case when ctx.fileType = FILE_XLSB then '.bin' else '.xml' end;
    tab.cols := CT_TableColumns();
    tab.cols.extend(columnMap.count);
    for i in 1 .. columnMap.count loop
      tab.cols(i).id := i;
      tab.cols(i).name := columnMap(i).name;
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
    wb.refStyle := ExcelFmla.REF_A1;
    return wb;
  end;
  
  procedure addDefaultStyles (
    styles  in out nocopy CT_Stylesheet
  )
  is
    styleXfId    pls_integer;
    defaultFont  CT_Font := styles.fonts(0);
    hlinkFont    CT_Font;
  begin
    if styles.hasHlink then
      -- new hyperlink font derived from default
      hlinkFont := makeFont(defaultFont.name, defaultFont.sz, defaultFont.b, defaultFont.i, 'theme:10', 'single');
      -- new master cell xf using this font
      styleXfId := putCellStyleXf(styles, makeCellXf(styles, null, font => hlinkFont)); -- master cell xf
      styles.hlinkXfId := styleXfId;
      -- new named cell style for builtinId 8 (= hyperlink style)
      putNamedCellStyle(styles, 'Hyperlink', styleXfId, 8);
    end if;
  end;

  procedure createStylesheet (
    ctx       in out nocopy context_t
  , styles    in CT_Stylesheet
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
    if styles.dxfs.count != 0 then
      stream_write(stream, '<dxfs count="'||to_char(styles.dxfs.count)||'">');
      for i in styles.dxfs.first .. styles.dxfs.last loop
        stream_write(stream, '<dxf>' || styles.dxfs(i).content || '</dxf>');
      end loop;
      stream_write(stream, '</dxfs>');      
    else
      stream_write(stream, '<dxfs count="0"/>');
    end if;
    
    stream_write(stream, '</styleSheet>');
    stream_flush(stream);
    
    addPart(ctx, 'xl/styles.xml', MT_STYLES, stream.content);
    
  end;

  procedure createStylesheetBin (
    ctx       in out nocopy context_t
  , styles    in CT_Stylesheet
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
        xutl_xlsb.put_Fill(stream, styles.fills(fillId));
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
    for xfId in styles.cellStyleXfs.first .. styles.cellStyleXfs.last loop
      -- BrtXF
      xutl_xlsb.put_XF(stream 
                     --, xfId       => styles.cellXfs(xfId).xfId
                     , numFmtId  => styles.cellStyleXfs(xfId).numFmtId
                     , fontId    => styles.cellStyleXfs(xfId).fontId
                     , fillId    => styles.cellStyleXfs(xfId).fillId
                     , borderId  => styles.cellStyleXfs(xfId).borderId
                     , alignment => styles.cellStyleXfs(xfId).alignment
                     );
    end loop;
    xutl_xlsb.put_simple_record(stream, 627);  -- BrtEndCellStyleXFs
    
    -- cellXfs
    if styles.cellXfs.count != 0 then
      xutl_xlsb.put_simple_record(stream, 617, int2raw(styles.cellXfs.count));  -- BrtBeginCellXFs
      for xfId in styles.cellXfs.first .. styles.cellXfs.last loop
        -- BrtXF
        xutl_xlsb.put_XF(stream 
                       , xfId      => styles.cellXfs(xfId).xfId
                       , numFmtId  => styles.cellXfs(xfId).numFmtId
                       , fontId    => styles.cellXfs(xfId).fontId
                       , fillId    => styles.cellXfs(xfId).fillId
                       , borderId  => styles.cellXfs(xfId).borderId
                       , alignment => styles.cellXfs(xfId).alignment
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
    if styles.dxfs.count != 0 then
      xutl_xlsb.put_simple_record(stream, 505, int2raw(styles.dxfs.count));  -- BrtBeginDXFs
      for i in styles.dxfs.first .. styles.dxfs.last loop
        xutl_xlsb.put_DXF(stream, styles.dxfs(i).style);
      end loop;    
      xutl_xlsb.put_simple_record(stream, 506);  -- BrtEndDXFs
    end if;
    
    -- tableStyles?
    
    xutl_xlsb.put_simple_record(stream, 279); -- BrtEndStyleSheet
    xutl_xlsb.flush_stream(stream);
    addPart(ctx, 'xl/styles.bin', MT_STYLES_BIN, stream.content);
    
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
        stream_write(stream, '<si>');
        if not ctx.string_list(i).isRichText then
          stream_write(stream, '<t>');
          stream_write(stream, ctx.string_list(i).value, escape_xml => true);
          stream_write(stream, '</t>');
        else
          stream_write(stream, ctx.string_list(i).value);
        end if;
        stream_write(stream, '</si>');
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
    stream       xutl_xlsb.Stream_T;
    textRuns     ExcelTypes.CT_TextRunList;
    strRunArray  xutl_xlsb.StrRunArray_T;
    str          varchar2(32767);
    pos          pls_integer;
  begin
    if ctx.string_cnt != 0 then
      stream := xutl_xlsb.new_stream();
      xutl_xlsb.put_BeginSst(stream, ctx.string_cnt, ctx.string_map.count); -- BrtBeginSst
      for i in 1 .. ctx.string_list.count loop
      
        if not ctx.string_list(i).isRichText then
          
          xutl_xlsb.put_SSTItem(stream, ctx.string_list(i).value);
          
        else
          
          textRuns := ctx.rt_cache(i).runs;
          strRunArray := xutl_xlsb.StrRunArray_T();
          str := null;
          pos := 0;
          for j in 1 .. textRuns.count loop
            strRunArray.extend;
            strRunArray(j).ich := pos;
            strRunArray(j).ifnt := putFont(ctx.workbook.styles, textRuns(j).font);
            str := str || textRuns(j).text;
            pos := pos + length(textRuns(j).text);
          end loop;
          
          xutl_xlsb.put_SSTItem(stream, str, strRunArray);
          
        end if;
      
      end loop;
      xutl_xlsb.put_simple_record(stream, 160); -- BrtEndSst
      xutl_xlsb.flush_stream(stream);
      addPart(ctx, 'xl/sharedStrings.bin', MT_SHAREDSTRINGS_BIN, stream.content);
    end if;
  end;
  
  function createImagePart (
    ctx      in out nocopy context_t
  , content  in blob
  )
  return image_t
  is
    p         integer;
    buf       raw(32767);
    fileSize  integer;
    sz        integer;

    image     image_t;
    ext       varchar2(16);
    checksum  varchar2(32);
    bucket    imageList_t;
    partIdx   pls_integer;
    
    rootQName   varchar2(512);
    svgContent  xmltype;
    svgWidth    varchar2(256);
    svgHeight   varchar2(256);
    
  begin
        
    -- PNG signature
    if dbms_lob.substr(content, 8) = '89504E470D0A1A0A' then   
      image.mimeType := MT_PNG;

      -- first chunk must be IHDR
      p := 17; -- skip signature [8] + IHDR (length+type) [8]
      image.width := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 4, p), utl_raw.big_endian); 
      image.height := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 4, p+4), utl_raw.big_endian);
      
      
    -- JPEG signature, just checking the first 3 bytes (SOI + 1st byte of next marker)
    elsif dbms_lob.substr(content, 3) = 'FFD8FF' then
      image.mimeType := MT_JPEG;
      
      p := 3; -- skip SOI
      fileSize := dbms_lob.getlength(content);
      
      while p < fileSize loop
        buf := dbms_lob.substr(content, 2, p); -- JFIF marker type
        if buf in ('FFC0','FFC1','FFC2','FFC3','FFC5','FFC6','FFC7','FFC8','FFC9','FFCA','FFCB','FFCD','FFCE','FFCF') then --SOFn
          image.height := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 2, p+5), utl_raw.big_endian);
          image.width := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 2, p+7), utl_raw.big_endian);
          exit;
        elsif buf = 'FFDA' then -- SOS
          exit;
        else
          sz := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 2, p+2), utl_raw.big_endian); -- segment size
        end if;
        p := p + sz + 2;
      end loop;
      
    
    -- GIF87a/89a signature, just checking the first 4 bytes (GIF8)
    elsif dbms_lob.substr(content, 4) = '47494638' then
      image.mimeType := MT_GIF;
      
      p := 7; -- skip signature [6]
      image.width := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 2, p), utl_raw.little_endian);
      image.height := utl_raw.cast_to_binary_integer(dbms_lob.substr(content, 2, p+2), utl_raw.little_endian);
      
      
    else
      
      -- check for SVG content
      begin
        svgContent := blobToXml(content);
        
        select '{'||nsUri||'}'||localName
        into rootQName
        from xmltable( '/*'
               passing svgContent
               columns localName varchar2(256) path 'local-name()'
                     , nsUri     varchar2(256) path 'namespace-uri()'
             );
              
        if rootQName != '{http://www.w3.org/2000/svg}svg' then
          error('Unsupported XML image type');
        end if;
        
        image.mimeType := MT_SVG;
        
        select width, height
        into svgWidth, svgHeight
        from xmltable(xmlnamespaces(default 'http://www.w3.org/2000/svg'), '/svg'
               passing svgContent
               columns width  varchar2(256) path '@width'
                     , height varchar2(256) path '@height'
             );
             
        image.width := ExcelTypes.convertToPx(svgWidth);
        image.height := ExcelTypes.convertToPx(svgHeight);
        
        if image.width is null or image.height is null then
          error('Missing SVG dimensions');
        end if;
             
      exception
        when xml_parse_exception then
          error('Unsupported image type');
      end;
      
    end if;
    
    ext := putDefaultExtension(ctx, image.mimeType);

    -- BLOB dedup?
    if hash_available then
      checksum := rawtohex(hashMD5(dbms_lob.substr(content, 2000))); -- MD5 of first 2000 bytes
      if ctx.imageHashMap.exists(checksum) then
        bucket := ctx.imageHashMap(checksum);
        for i in 1 .. bucket.count loop
          -- iterate through parts referenced in the matching bucket
          partIdx := ctx.pck.partIndices(bucket(i).partName);
          if dbms_lob.compare(content, ctx.pck.parts(partIdx).contentBin) = 0 then
            image := bucket(i);
            exit;
          end if;
        end loop;
      end if;
    end if;
  
    if image.partName is null then
      
      ctx.imageNameSeq := ctx.imageNameSeq + 1;
      image.partName := 'xl/media/image'||to_char(ctx.imageNameSeq)||'.'||ext;
      image.name := 'Image '||to_char(ctx.imageNameSeq);
      
      -- add image checksum if supported
      if checksum is not null then
        if bucket is null then
          bucket := imageList_t();
        end if;
        bucket.extend;
        bucket(bucket.last) := image;
        ctx.imageHashMap(checksum) := bucket;
      end if;
      
      addPart(ctx, image.partName, null, content);
    
    end if;
    
    return image;
  end;

  function createFallbackImage (
    width   in pls_integer
  , height  in pls_integer
  )
  return blob
  is

    type chunk_t is record (ptr integer, len integer);
    chnk       chunk_t;

    output     blob;
    input      blob;
    p          integer := 1;
    gz         blob;
    
    bitDepth   pls_integer := 1;    
    pattern1   raw(1024);
    pattern2   raw(1024);
    
    scanline1  raw(32767);
    scanline2  raw(32767);
    numPacks   pls_integer;

    procedure write (bytes in raw)
    is
      len  pls_integer := utl_raw.length(bytes);
    begin
      dbms_lob.writeappend(output, len, bytes);
      p := p + len;
    end;
    
    procedure beginChunk (chunkType in varchar2) 
    is
    begin
      chnk.ptr := p;
      chnk.len := 0;
      write('00000000'); -- length placeholder
      write(utl_raw.cast_to_raw(chunkType));
    end;
    
    procedure writeChunkData (bytes in raw) is
    begin
      write(bytes);
      chnk.len := chnk.len + utl_raw.length(bytes);
    end;
    
    procedure endChunk is
    begin
      dbms_lob.write(output, 4, chnk.ptr, utl_raw.cast_from_binary_integer(chnk.len));
      write(crc32(output, chnk.ptr + 4, chnk.len + 4));
    end;

  begin
        
    dbms_lob.createtemporary(output, true);
    
    write('89504E470D0A1A0A'); -- PNG signature
    
    beginChunk('IHDR');
    writeChunkData(utl_raw.cast_from_binary_integer(width));
    writeChunkData(utl_raw.cast_from_binary_integer(height));
    writeChunkData(utl_raw.substr(utl_raw.cast_from_binary_integer(bitDepth, utl_raw.little_endian),1,1)); -- bit depth
    writeChunkData('03');     -- color type (palette)
    writeChunkData('000000'); -- CM (0) + filter (0) + interlace (no)
    endChunk();
    
    beginChunk('PLTE');
    writeChunkData('E0E0E0'); -- color 0 (lightgray)
    writeChunkData('FFFFFF'); -- color 1 (white)
    endChunk();

    beginChunk('IDAT');
    -- begin ZLIB content
    writeChunkData('789C'); -- zlib header with default compression flags
    pattern1 := '00000000FFFFFFFF'; -- 32px of color 0 + 32px of color 1
    pattern2 := 'FFFFFFFF00000000'; -- 32px of color 1 + 32px of color 0
    
    numPacks := ceil(width/64); -- number of pattern units
    scanline1 := utl_raw.concat('00', utl_raw.substr(utl_raw.copies(pattern1, numPacks), 1, ceil(bitDepth*width/8)));
    scanline2 := utl_raw.concat('00', utl_raw.substr(utl_raw.copies(pattern2, numPacks), 1, ceil(bitDepth*width/8)));
    
    dbms_lob.createtemporary(input, true);
    for i in 0 .. height-1 loop
      dbms_lob.writeappend(input, utl_raw.length(scanline1), case when mod(floor(i/32),2) = 0 then scanline1 else scanline2 end);
    end loop;
    
    gz := utl_compress.lz_compress(input);
      
    writeChunkData(dbms_lob.substr(gz, dbms_lob.getlength(gz)-18, 11)); -- extract DEFLATE content
    writeChunkData(adler32(input));
    -- end ZLIB content
    endChunk();
    
    dbms_lob.freetemporary(input);
    
    beginChunk('IEND');
    endChunk();
    
    return output;

  end;  

  function createImageMetadata (
    ctx      in out nocopy context_t
  , content  in blob
  )
  return pls_integer
  is
    image  image_t := createImagePart(ctx, content);
    vmId   pls_integer;
  begin
    
    if ctx.imageMetadata.vmMap.exists(image.partName) then
      vmId := ctx.imageMetadata.vmMap(image.partName);
    else
      vmId := nvl(ctx.imageMetadata.partMap.last, 0) + 1;
      ctx.imageMetadata.vmMap(image.partName) := vmId;
      ctx.imageMetadata.partMap(vmId) := image.partName;
    end if;
    
    return vmId;
  end;
  
  function createDrawingPart (
    ctx  in out nocopy context_t
  , dr   in CT_Drawing
  )
  return varchar2
  is
    stream  stream_t := new_stream();
    image   image_t;
    image2  image_t; -- fallback image
    part    part_t;
    rId     varchar2(256);
    rId2    varchar2(256); -- fallback rId
    a       CT_DrawingAnchor;
    
    procedure write_marker (tag in varchar2, mk in CT_Marker) is
    begin
      stream_write(stream, '<'||tag||'>');
      stream_write(stream, '<col>'||mk.col||'</col><colOff>'||mk.colOff||'</colOff><row>'||mk."row"||'</row><rowOff>'||mk.rowOff||'</rowOff>');
      stream_write(stream, '</'||tag||'>');
    end;

    procedure write_pos is
    begin
      stream_write(stream, '<pos x="'||a.pos.x||'" y="'||a.pos.y||'"/>');
    end;
    
    procedure write_ext is
    begin
      stream_write(stream, '<ext cx="'||nvl(a.ext.cx, to_char(image.width)||'px')||'" cy="'||nvl(a.ext.cy, to_char(image.height)||'px')||'"/>');
    end;
    
    procedure write_pic_info is
    begin
      stream_write(stream, '<pic>');
      stream_write(stream, '<nvPicPr><cNvPr id="'||to_char(a.id)||'" name="'||dbms_xmlgen.convert(image.name)||'"/><cNvPicPr><a:picLocks noChangeAspect="1"/></cNvPicPr></nvPicPr>');    
      stream_write(stream, '<blipFill><a:blip r:embed="'||rId||'">');
      if rId2 is not null then
        stream_write(stream, '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="'||rId2||'"/></a:ext></a:extLst>');
      end if;
      stream_write(stream, '</a:blip><a:stretch><a:fillRect/></a:stretch></blipFill>');
      stream_write(stream, '<spPr><a:prstGeom prst="rect"/></spPr>');
      stream_write(stream, '</pic>');
      stream_write(stream, '<clientData/>');      
    end;
    
  begin
    
    ctx.drawingIdSeq := ctx.drawingIdSeq + 1;
    part := new_part('xl/drawings/drawing'||to_char(ctx.drawingIdSeq)||'.xml', MT_DRAWING);
  
    stream_write(stream, '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    
    for i in 1 .. dr.count loop
      
      a := dr(i);
      image := createImagePart(ctx, a.content);
      rId := addRelationship(part, RS_IMAGE, image.partName);
      rId2 := null;
      
      -- If the image is SVG, generate a dummy fallback image in PNG format to be displayed by incompatible Office versions
      if image.mimeType = MT_SVG then
        rId2 := rId;
        image2 := createImagePart(ctx, createFallbackImage(image.width, image.height));
        rId := addRelationship(part, RS_IMAGE, image2.partName);
      end if;
    
      case a.anchorType
      when TWOCELL_ANCHOR then
        
        stream_write(stream, '<twoCellAnchor editAs="');
        stream_write(stream, case a.editAs
                             when MOVE_RESIZE then 'twoCell'
                             when MOVE_NO_RESIZE then 'oneCell'
                             when NO_MOVE_NO_RESIZE then 'absolute'
                             end);
        stream_write(stream, '">'); 
        write_marker('from', a."from");
        write_marker('to', a."to");
        write_pic_info(); 
        stream_write(stream, '</twoCellAnchor>');
      
      when ONECELL_ANCHOR then
        
        stream_write(stream, '<oneCellAnchor>'); 
        write_marker('from', a."from");
        write_ext();
        write_pic_info(); 
        stream_write(stream, '</oneCellAnchor>');    
      
      when ABSOLUTE_ANCHOR then
        
        stream_write(stream, '<absoluteAnchor>'); 
        write_pos();
        write_ext();
        write_pic_info(); 
        stream_write(stream, '</absoluteAnchor>');
      
      end case;
    
    end loop;
    
    stream_write(stream, '</wsDr>');
    stream_flush(stream);
    part.content := stream.content;
    
    addPart(ctx, part);
    
    return part.name;
  end;

  procedure writeRowStart (
    stream  in out nocopy stream_t
  , r       in row_t
  )
  is
  begin
    stream_write(stream, '<row r="'||to_char(r.id)||'"'
                        ||case when r.props.xfId is not null then ' s="'||to_char(r.props.xfId)||'" customFormat="1"' end
                        ||case when r.props.height is not null then ' ht="'||to_char(r.props.height, 'TM9', NLS_PARAM_STRING)||'" customHeight="1"' end
                        ||'>');
  end;

  procedure writeRowBin (
    stream         in out nocopy xutl_xlsb.stream_t
  , r              in row_t
  , defaultHeight  in number
  )
  is
  begin
    xutl_xlsb.put_RowHdr(stream
                       , rowIndex => r.id - 1
                       , height   => r.props.height
                       , styleRef => r.props.xfId
                       , defaultHeight => defaultHeight
                       );
  end;
  
  procedure writeCell (
    ctx     in out nocopy context_t
  , stream  in out nocopy stream_t
  , cell    in cell_t
  )
  is
    cellRef  varchar2(10) := cell.c||to_char(cell.r);
    sst_idx  pls_integer;
    vmId     pls_integer;
  begin

    case cell.v.st
    when ST_STRING then
      if cell.v.varchar2_value is not null then
        sst_idx := put_string(ctx, cell.v.varchar2_value);
        stream_write(stream, '<c r="'||cellRef
              ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
              ||'" t="s"><v>'||to_char(sst_idx - 1)||'</v></c>');
      else
        stream_write(stream, '<c r="'||cellRef
              ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
              ||'"></c>');
      end if;
              
    when ST_NUMBER then
      stream_write(stream, '<c r="'||cellRef
          ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
          ||'"><v>'||cell.v.varchar2_value||'</v></c>');
              
    when ST_DATETIME then
      stream_write(stream, '<c r="'||cellRef||'" s="'||to_char(cell.xfId)||'"><v>'||cell.v.varchar2_value||'</v></c>');
              
    when ST_LOB then
      if cell.v.clob_value is not null and dbms_lob.getlength(cell.v.clob_value) != 0 then
        -- try conversion to VARCHAR2
        begin
          sst_idx := put_string(ctx, stripXmlControlChars(to_char(cell.v.clob_value)));
          stream_write(stream, '<c r="'||cellRef
              ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
              ||'" t="s"><v>'||to_char(sst_idx - 1)||'</v></c>');
        exception
          when value_error then
            -- stream CLOB content as inlineStr, up to 32767 chars
            stream_write(stream, '<c r="'||cellRef
                ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
                ||'" t="inlineStr"><is><t>');
            stream_write_clob(stream, cell.v.clob_value, 32767, true);
            stream_write(stream, '</t></is></c>');
        end;
      end if;
      
    when ST_RICHTEXT then
      sst_idx := put_rt(ctx, ExcelTypes.makeRichText(cell.v.xml_value, getCellFont(ctx, cell.xfId)));
      stream_write(stream, '<c r="'||cellRef
            ||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end
            ||'" t="s"><v>'||to_char(sst_idx - 1)||'</v></c>');
              
    when ST_FORMULA then
      stream_write(stream, '<c r="'||cellRef||case when cell.xfId != 0 then '" s="'||to_char(cell.xfId) end||'">'||
                           case when cell.f.shared then
                             '<f t="shared" si="'||to_char(cell.f.sharedIdx)||'"' ||
                             case when cell.f.hasRef then ' ref="###'||to_char(cell.f.sharedIdx)||'###">' ||
                               dbms_xmlgen.convert(ExcelFmla.parse(cell.f.expr, p_cellRef => cellRef, p_refStyle => cell.f.refStyle)) ||
                               '</f>' 
                             else
                               '/>'
                             end
                           else
                             '<f>'||dbms_xmlgen.convert(ExcelFmla.parse(cell.f.expr, p_cellRef => cellRef, p_refStyle => cell.f.refStyle))||'</f>'
                           end ||
                           '</c>');
    
    when ST_IMAGE then
      vmId := createImageMetadata(ctx, cell.v.blob_value);
      stream_write(stream, '<c r="'||cellRef||'"'
          ||case when cell.xfId != 0 then ' s="'||to_char(cell.xfId)||'"' end
          ||' t="e" vm="'||to_char(vmId)||'"><v>#VALUE!</v></c>');
      
    end case;
        
  end;

  procedure writeCellBin (
    ctx     in out nocopy context_t
  , stream  in out nocopy xutl_xlsb.stream_t
  , cell    in cell_t
  )
  is
    sst_idx  pls_integer;
    vmId     pls_integer;
  begin

    case cell.v.st
    when ST_STRING then
      if cell.v.varchar2_value is not null then
        sst_idx := put_string(ctx, cell.v.varchar2_value);
        xutl_xlsb.put_CellIsst(stream, cell.cn-1, cell.xfId, sst_idx-1);
      else
        -- put a blank cell
        xutl_xlsb.put_CellNumber(stream, cell.cn-1, cell.xfId, null);
      end if;
              
    when ST_NUMBER then
      xutl_xlsb.put_CellNumber(stream, cell.cn-1, cell.xfId, cell.v.number_value);
              
    when ST_DATETIME then
      xutl_xlsb.put_CellNumber(stream, cell.cn-1, cell.xfId, cell.v.number_value);
              
    when ST_LOB then      
      if cell.v.clob_value is not null and dbms_lob.getlength(cell.v.clob_value) != 0 then
        -- try conversion to VARCHAR2
        begin
          sst_idx := put_string(ctx, to_char(cell.v.clob_value));
          xutl_xlsb.put_CellIsst(stream, cell.cn-1, cell.xfId, sst_idx-1);
        exception
          when value_error then
            -- stream CLOB content as an inline string, up to 32767 chars
            xutl_xlsb.put_CellSt(stream, cell.cn-1, cell.xfId, lobValue => cell.v.clob_value);
        end;
      end if;
      
    when ST_RICHTEXT then
      sst_idx := put_rt(ctx, ExcelTypes.makeRichText(cell.v.xml_value, getCellFont(ctx, cell.xfId)));
      xutl_xlsb.put_CellIsst(stream, cell.cn-1, cell.xfId, sst_idx-1);
    
    when ST_FORMULA then
      xutl_xlsb.put_CellFmla(
        stream
      , colIndex => cell.cn-1
      , styleRef => cell.xfId
      , expr     => cell.f.expr
      , shared   => cell.f.shared
      , si       => cell.f.sharedIdx
      , cellRef  => cell.c||to_char(cell.r)
      , refStyle => cell.f.refStyle
      );
      
    when ST_IMAGE then
      vmId := createImageMetadata(ctx, cell.v.blob_value);
      xutl_xlsb.put_CellImage(
        stream
      , colIndex => cell.cn-1
      , styleRef => cell.xfId
      , vmId     => vmId
      ); 
    
    end case;
        
  end;

  procedure setAnchorRowOffset (sd in sheet_definition_t, anchorRef in out nocopy anchorRef_t) is
    anchorTable         table_t;
    anchorTableRowSpan  pls_integer := 0;
  begin
    if anchorRef.tableId is not null then
      anchorTable := sd.tableList(anchorRef.tableId);
      if anchorRef.anchorPosition in (BOTTOM_LEFT,BOTTOM_RIGHT) then
        if not anchorTable.range.is_null then
          anchorTableRowSpan := anchorTable.range.end_ref.r - anchorTable.range.start_ref.r + 1;
        end if;
        anchorRef.rowOffset := anchorTable.anchorRef.rowOffset + anchorTableRowSpan - 1 + anchorRef.rowOffset;
      elsif anchorRef.anchorPosition in (TOP_LEFT,TOP_RIGHT) then
        anchorRef.rowOffset := anchorTable.anchorRef.rowOffset + anchorRef.rowOffset;
      end if;
    end if;
  end;

  procedure setAnchorColOffset (sd in sheet_definition_t, anchorRef in out nocopy anchorRef_t) is
    anchorTable         table_t;
    anchorTableColSpan  pls_integer;
  begin
    if anchorRef.tableId is not null then
      anchorTable := sd.tableList(anchorRef.tableId);
      if anchorRef.anchorPosition in (TOP_RIGHT,BOTTOM_RIGHT) then
        anchorTableColSpan := anchorTable.sqlMetadata.visibleColumnSet.count;
        anchorRef.colOffset := anchorTable.anchorRef.colOffset + anchorTableColSpan - 1 + anchorRef.colOffset;
      elsif anchorRef.anchorPosition in (TOP_LEFT,BOTTOM_LEFT) then
        anchorRef.colOffset := anchorTable.anchorRef.colOffset + anchorRef.colOffset;
      end if;
    end if;
  end;

  procedure applyRangeStyles (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t    
  )
  is
    cellSpan          cellSpan_t;
    defaultRangeXfId  pls_integer;
    rangeXfId         pls_integer;
    cell              cell_t;
  begin
    for i in 1 .. sd.cellRanges.count loop
      cellSpan := sd.cellRanges(i).span;
      setAnchorRowOffset(sd, cellSpan.anchorRef);
      setAnchorColOffset(sd, cellSpan.anchorRef);
      
      defaultRangeXfId := sd.cellRanges(i).xfId;
      if sd.defaultXfId is not null then
        defaultRangeXfId := mergeCellStyle(ctx, sd.defaultXfId, defaultRangeXfId);
      end if;
      
      for rowIdx in cellSpan.anchorRef.rowOffset .. cellSpan.anchorRef.rowOffset + cellSpan.rowSpan - 1 loop      
        for colIdx in cellSpan.anchorRef.colOffset .. cellSpan.anchorRef.colOffset + cellSpan.colSpan - 1 loop
          
          rangeXfId := defaultRangeXfId;
          
          if sd.cellRanges(i).outsideBorders then
            rangeXfId := setRangeBorders(rangeXfId, cellSpan, rowIdx, colIdx);
          end if;
        
          if sd.data.rows.exists(rowIdx) and sd.data.rows(rowIdx).cells.exists(colIdx) then
            cell := sd.data.rows(rowIdx).cells(colIdx);
            cell.xfId := mergeCellStyle(ctx, rangeXfId, cell.xfId);
          else
            cell := null;
            cell.r := rowIdx;
            cell.cn := colIdx;
            cell.c := base26encode(cell.cn);
            cell.v.st := ST_NUMBER;
            cell.xfId := rangeXfId;
            cell.isTableCell := false;
          end if;
          sd.data.rows(rowIdx).cells(colIdx) := cell;
        
        end loop;
      end loop;
    end loop;
  end;

  procedure validateName (input in varchar2) is
    tmp  varchar2(32767) := upper(input);
  begin
    if length(tmp) > 255 then
      error('Defined name is too long: %s', input);
    end if;
    if ExcelFmla.isValidCellReference(tmp)
      or tmp in ('R','C','RC') 
      or regexp_like(tmp, '^C0*[1-9]')
    then
      error('Invalid defined name: %s', input);
    end if;
    tmp := regexp_substr(tmp, '^R0*([1-9]\d*)', 1, 1, null, 1);
    if tmp is not null and length(tmp) < 8 and to_number(tmp) between 1 and MAX_ROW_NUMBER then
      error('Invalid defined name: %s', input);
    end if;
  end;

  procedure putNameImpl (
    ctxId       in ctxHandle
  , name        in varchar2
  , value       in varchar2
  , sheetName   in varchar2
  , cellRef     in varchar2 default null
  , comment     in varchar2 default null
  , hidden      in boolean default false
  , futureFunc  in boolean default false
  , builtIn     in boolean default false
  , refStyle    in pls_integer default null
  )
  is
    nameKey      varchar2(2048);
    definedName  ExcelTypes.CT_DefinedName;
  begin
    loadContext(ctxId);
    
    definedName.scope := sheetName;
    
    nameKey := upper(case when definedName.scope is not null then definedName.scope || '!' end || name);
    if currentCtx.nameMap.exists(nameKey) then
      error('Defined name already exists: %s', name);
    end if;
    
    validateName(name);
    
    currentCtx.names.extend;
    definedName.idx := currentCtx.names.last;
    definedName.name := name;
    
    definedName.formula := value;
    definedName.cellRef := cellRef;
    definedName.refStyle := refStyle;
    definedName.comment := comment;
    definedName.hidden := nvl(hidden, false);
    definedName.futureFunction := nvl(futureFunc, false);
    definedName.builtIn := nvl(builtIn, false);
    
    currentCtx.names(definedName.idx) := definedName;
    currentCtx.nameMap(nameKey) := definedName;
    
    if definedName.builtIn then
      ExcelFmla.addDefinedName(definedName);
    end if;
  
  end;

  procedure setSqref (
    sqref     in out nocopy ExcelTypes.cellRangeSeq_t
  , cellRange in ExcelTypes.ST_Sqref
  )
  is
    dvRange   ExcelTypes.cellRange_t;
    tmpRange  range_t;
    
    boundingAreaFirstRw   pls_integer;
    boundingAreaFirstCol  pls_integer;
  begin
    
    if cellRange is not null and cellRange.count != 0 then
      
      sqref.ranges := ExcelTypes.cellRangeList_t();
      for i in 1 .. cellRange.count loop
        dvRange.value := cellRange(i);
        tmpRange := parseRangeExpr(dvRange.value);
        dvRange.rwFirst := tmpRange.start_ref.r;
        dvRange.rwLast := nvl(tmpRange.end_ref.r, tmpRange.start_ref.r);
        dvRange.colFirst := tmpRange.start_ref.cn;
        dvRange.colLast := nvl(tmpRange.end_ref.cn, tmpRange.start_ref.cn);
        sqref.ranges.extend;
        sqref.ranges(sqref.ranges.last) := dvRange;
        
        if i = 1 then
          boundingAreaFirstRw := dvRange.rwFirst;
          boundingAreaFirstCol := dvRange.colFirst;
        else
          boundingAreaFirstRw := least(boundingAreaFirstRw, dvRange.rwFirst);
          boundingAreaFirstCol := least(boundingAreaFirstCol, dvRange.colFirst);
        end if;
        
      end loop;
      
      -- Data Validation context:
      -- active cell reference, used as a point of origin to resolve relative references occurring in fmla1 and fmla2
      -- by Excel convention, it is the top-left cell of the last range in the sequence
      sqref.lastRangeCellRef := tmpRange.start_ref.value;
      
      -- Data Validation context:
      -- top-left cell reference of the rectangular bounding area of all ranges in the sequence
      -- in the XLSX format, this is used as new point of origin to resolve relative references, which are also translated accordingly
      -- in the XLSB format, this is not used, an offset is stored instead
      sqref.boundingAreaFirstCellRef := makeCellRef(boundingAreaFirstCol, boundingAreaFirstRw);
      
    end if;
    
  end;
  
  function makeDVRule (
    p_type              in varchar2
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
  )
  return ExcelTypes.CT_DataValidation
  is
    dvRule  ExcelTypes.CT_DataValidation;
  begin
    
    if not ExcelTypes.isValidDataValidationType(p_type) then
      error('Invalid data validation type: %s', p_type);
    end if;
    dvRule.type := p_type;
    
    setSqref(dvRule.sqref, p_cellRange);
    
    if dvRule.type not in ('list','custom') then
      if not ExcelTypes.isValidDataValidationOperator(p_operator) then
        error('Invalid data validation operator: %s', p_operator);
      end if;
      dvRule.operator := p_operator;
    else
      dvRule.operator := 'between'; -- default value, useful for XLSB format
    end if;
    
    if p_value1 is null then
      error('Missing value (%s) in data validation rule', 
        case 
        when dvRule.type in ('whole','decimal','textLength') then
          case 
          when dvRule.operator in ('between','notBetween','greaterThan','greaterThanOrEqual') then 'Minimum'
          when dvRule.operator in ('lessThan', 'lessThanOrEqual') then 'Maximum'
          else case when dvRule.type = 'textLength' then 'Length' else 'Value' end
          end
        when dvRule.type in ('date','time') then
          case 
          when dvRule.operator in ('between','notBetween','greaterThan','greaterThanOrEqual') then 'Start '||dvRule.type
          when dvRule.operator in ('lessThan', 'lessThanOrEqual') then 'End '||dvRule.type
          else initcap(dvRule.type)
          end
        when dvRule.type = 'list' then 'Source'
        else 'Formula'
        end
      );
    end if;
    
    dvRule.fmla1 := p_value1;
    
    if dvRule.operator in ('between','notBetween') and dvRule.type not in ('list','custom') then
      if p_value2 is null then
        error('Missing value (%s) in data validation rule',
          case 
          when dvRule.type in ('whole','decimal','textLength') then 'Maximum'
          else 'End '||dvRule.type
          end
        );
      end if;
      dvRule.fmla2 := p_value2;
    end if;
    
    dvRule.allowBlank := nvl(p_allowBlank, true);
    dvRule.showDropDown := nvl(p_showDropDown, true);
    
    dvRule.showErrorMessage := nvl(p_showErrorMessage, true);
    dvRule.errorStyle := nvl(p_errorStyle, 'stop');
    if p_errorMsg is not null or p_errorTitle is not null then   
      dvRule.error := p_errorMsg;
      dvRule.errorTitle := p_errorTitle;
      if p_errorStyle is not null and not ExcelTypes.isValidDataValidationErrStyle(p_errorStyle) then
          error('Invalid data validation error style: %s', p_errorStyle);
      end if;
    end if;

    dvRule.showInputMessage := nvl(p_showInputMessage, true);
    if p_promptMsg is not null or p_promptTitle is not null then
      dvRule.prompt := p_promptMsg;
      dvRule.promptTitle := p_promptTitle;
    end if;
    
    dvRule.refStyle1 := p_refStyle1;
    dvRule.refStyle2 := p_refStyle2;
    
    return dvRule;
    
  end;

  function makeCfRule (
    p_type        in pls_integer
  , p_cellRange   in ExcelTypes.ST_Sqref
  , p_style       in ExcelGen.cfmtStyleHandle
  , p_operator    in pls_integer
  , p_value1      in varchar2
  , p_value2      in varchar2
  , p_param       in pls_integer
  , p_percent     in boolean
  , p_cfvoList    in ExcelTypes.CT_CfvoList
  , p_hideValue   in boolean
  , p_iconSet     in pls_integer
  , p_reverse     in boolean
  , p_stopIfTrue  in boolean
  , p_refStyle1   in pls_integer
  , p_refStyle2   in pls_integer
  )
  return ExcelTypes.CT_CfRule
  is
    rule               ExcelTypes.CT_CfRule;
    requiredCfvoCount  pls_integer;
    
    -- update cell ref placeholder with the top-left cell of the bounding area of all ranges this rule applies to.
    procedure updateCellRefPlaceholder is
    begin
      rule.fmla1 := replace(rule.fmla1, '__A1__', rule.sqref.boundingAreaFirstCellRef);      
    end;
    
  begin

    setSqref(rule.sqref, p_cellRange);
    rule.dxfId := p_style;
    rule.stopTrue := nvl(p_stopIfTrue, false);
    
    if not ExcelTypes.isValidCondFmtRuleType(p_type) then
      error('Invalid conditional formatting rule type');
    end if;

    case p_type
    when ExcelTypes.CF_TYPE_EXPR then
    
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_FMLA;
      rule.fmla1 := p_value1;
      rule.refStyle1 := p_refStyle1;
    
    when ExcelTypes.CF_TYPE_CELLIS then
      
      rule.type := ExcelTypes.CF_TYPE_CELLIS;
      rule.template := ExcelTypes.CF_TEMP_EXPR;
      
      if not ExcelTypes.isValidCondFmtOperator(p_type, p_operator) then
        error('Invalid relational operator for this conditional formatting rule type');
      end if;
      
      rule.param := p_operator; -- CF_OPER_*
      rule.fmla1 := p_value1;
      rule.refStyle1 := p_refStyle1;
      if rule.param in (ExcelTypes.CF_OPER_BN, ExcelTypes.CF_OPER_NB) then
        rule.fmla2 := p_value2;
        rule.refStyle2 := p_refStyle2;
      end if;
      
    when ExcelTypes.CF_TYPE_TEXT then
    
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_CONTAINSTEXT;

      if not ExcelTypes.isValidCondFmtOperator(p_type, p_operator) then
        error('Invalid text operator for this conditional formatting rule type');
      end if;
      
      case p_operator
      when ExcelTypes.CF_TEXTOPER_CONTAINS then
        rule.fmla1 := 'NOT(ISERROR(SEARCH("__TEXT__",__A1__)))';
        
      when ExcelTypes.CF_TEXTOPER_NOTCONTAINS then
        rule.fmla1 := 'ISERROR(SEARCH("__TEXT__",__A1__))';
        
      when ExcelTypes.CF_TEXTOPER_BEGINSWITH then
        rule.fmla1 := 'LEFT(__A1__,LEN("__TEXT__"))="__TEXT__"';
        
      when ExcelTypes.CF_TEXTOPER_ENDSWITH then
        rule.fmla1 := 'RIGHT(__A1__,LEN("__TEXT__"))="__TEXT__"';
        
      end case;
      
      updateCellRefPlaceholder;
      rule.fmla1 := replace(rule.fmla1, '__TEXT__', replace(p_value1, '"', '""'));
      
      rule.param := p_operator;
      rule.strParam := p_value1;
      
    when ExcelTypes.CF_TYPE_TIMEPERIOD then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;

      if not ExcelTypes.isValidCondFmtOperator(p_type, p_operator) then
        error('Invalid time period for this conditional formatting rule type');
      end if;

      case p_operator
      when ExcelTypes.CF_TIMEPERIOD_TODAY then
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODTODAY;
        rule.fmla1 := 'FLOOR(__A1__,1)=TODAY()';
        
      when ExcelTypes.CF_TIMEPERIOD_YESTERDAY then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODYESTERDAY;
        rule.fmla1 := 'FLOOR(__A1__,1)=TODAY()-1';
        
      when ExcelTypes.CF_TIMEPERIOD_LAST7DAYS then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODLAST7DAYS;
        rule.fmla1 := 'AND(TODAY()-FLOOR(__A1__,1)<=6,FLOOR(__A1__,1)<=TODAY())';
        
      when ExcelTypes.CF_TIMEPERIOD_THISWEEK then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODTHISWEEK;
        rule.fmla1 := 'AND(TODAY()-ROUNDDOWN(__A1__,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(__A1__,0)-TODAY()<=7-WEEKDAY(TODAY()))';
        
      when ExcelTypes.CF_TIMEPERIOD_LASTWEEK then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODLASTWEEK;
        rule.fmla1 := 'AND(TODAY()-ROUNDDOWN(__A1__,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(__A1__,0)<(WEEKDAY(TODAY())+7))';
        
      when ExcelTypes.CF_TIMEPERIOD_LASTMONTH then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODLASTMONTH;
        rule.fmla1 := 'AND(MONTH(__A1__)=MONTH(EDATE(TODAY(),0-1)),YEAR(__A1__)=YEAR(EDATE(TODAY(),0-1)))';
        
      when ExcelTypes.CF_TIMEPERIOD_TOMORROW then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODTOMORROW;
        rule.fmla1 := 'FLOOR(__A1__,1)=TODAY()+1';
        
      when ExcelTypes.CF_TIMEPERIOD_NEXTWEEK then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODNEXTWEEK;
        rule.fmla1 := 'AND(ROUNDDOWN(__A1__,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(__A1__,0)-TODAY()<(15-WEEKDAY(TODAY())))';
        
      when ExcelTypes.CF_TIMEPERIOD_NEXTMONTH then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODNEXTMONTH;
        rule.fmla1 := 'AND(MONTH(__A1__)=MONTH(EDATE(TODAY(),0+1)),YEAR(__A1__)=YEAR(EDATE(TODAY(),0+1)))';
        
      when ExcelTypes.CF_TIMEPERIOD_THISMONTH then 
        rule.template := ExcelTypes.CF_TEMP_TIMEPERIODTHISMONTH;
        rule.fmla1 := 'AND(MONTH(__A1__)=MONTH(TODAY()),YEAR(__A1__)=YEAR(TODAY()))';
        
      end case;
      
      updateCellRefPlaceholder;
      rule.param := p_operator;
    
    when ExcelTypes.CF_TYPE_BLANKS then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_CONTAINSBLANKS;
      rule.fmla1 := 'LEN(TRIM(__A1__))=0';
      updateCellRefPlaceholder;

    when ExcelTypes.CF_TYPE_NOBLANKS then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_CONTAINSNOBLANKS;
      rule.fmla1 := 'LEN(TRIM(__A1__))>0';
      updateCellRefPlaceholder;

    when ExcelTypes.CF_TYPE_ERRORS then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_CONTAINSERRORS;
      rule.fmla1 := 'ISERROR(__A1__)';
      updateCellRefPlaceholder;

    when ExcelTypes.CF_TYPE_NOERRORS then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_CONTAINSNOERRORS;
      rule.fmla1 := 'NOT(ISERROR(__A1__))';
      updateCellRefPlaceholder;
      
    when ExcelTypes.CF_TYPE_TOP then
      
      rule.type := ExcelTypes.CF_TYPE_FILTER;
      rule.template := ExcelTypes.CF_TEMP_FILTER;
      rule.percent := nvl(p_percent, false);
      if rule.percent and p_param not between 1 and 100 then
        error('Invalid rank percentage in conditional formatting rule');
      elsif p_param not between 1 and 1000 then
        error('Rank out of range in conditional formatting rule');
      end if;
      rule.param := p_param;
      rule.bottom := false;

    when ExcelTypes.CF_TYPE_BOTTOM then
      
      rule.type := ExcelTypes.CF_TYPE_FILTER;
      rule.template := ExcelTypes.CF_TEMP_FILTER;
      rule.percent := nvl(p_percent, false);
      if rule.percent and p_param not between 1 and 100 then
        error('Invalid rank percentage in conditional formatting rule');
      elsif p_param not between 1 and 1000 then
        error('Rank out of range in conditional formatting rule');
      end if;
      rule.param := p_param;
      rule.bottom := true;

    when ExcelTypes.CF_TYPE_EQUALABOVEAVERAGE then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_EQUALABOVEAVERAGE;
      rule.param := 0;
    
    when ExcelTypes.CF_TYPE_EQUALBELOWAVERAGE then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_EQUALBELOWAVERAGE;
      rule.param := 0;
    
    when ExcelTypes.CF_TYPE_ABOVEAVERAGE then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_ABOVEAVERAGE;
      rule.param := nvl(p_param, 0);
      if rule.param not between 0 and 3 then
        error('Invalid standard deviation specified in conditional formatting rule');
      end if;

    when ExcelTypes.CF_TYPE_BELOWAVERAGE then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_BELOWAVERAGE;
      rule.param := nvl(p_param, 0);
      if rule.param not between 0 and 3 then
        error('Invalid standard deviation specified in conditional formatting rule');
      end if;
      
    when ExcelTypes.CF_TYPE_DUPLICATES then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_DUPLICATEVALUES;

    when ExcelTypes.CF_TYPE_UNIQUES then
      
      rule.type := ExcelTypes.CF_TYPE_EXPRIS;
      rule.template := ExcelTypes.CF_TEMP_UNIQUEVALUES;
      
    when ExcelTypes.CF_TYPE_COLORSCALE then
      
      rule.type := ExcelTypes.CF_TYPE_GRADIENT;
      rule.template := ExcelTypes.CF_TEMP_GRADIENT;
      
      rule.cfvoList := p_cfvoList;
      
      if rule.cfvoList is null or rule.cfvoList.count = 0 then
        error('No value objects found for this color scale rule');
      elsif rule.cfvoList.count not in (2, 3) then
        error('Wrong number of value objects for this color scale rule');
      end if;
      
      for i in 1 .. rule.cfvoList.count loop
        if not ExcelTypes.isValidCondFmtVOType(rule.cfvoList(i).type) then
          error('Invalid value object type for this color scale rule');
        end if;
        if rule.cfvoList(i).color is null then
          error('Missing color for this color scale point');
        end if;
        rule.cfvoList(i).color := ExcelTypes.validateColor(rule.cfvoList(i).color);
      end loop;
      
    when ExcelTypes.CF_TYPE_DATABAR then
      
      rule.type := ExcelTypes.CF_TYPE_DATABAR;
      rule.template := ExcelTypes.CF_TEMP_DATABAR;
      
      rule.cfvoList := p_cfvoList;
      
      if rule.cfvoList is null or rule.cfvoList.count = 0 then
        error('No value objects found for this data bar rule');
      elsif rule.cfvoList.count != 3 then
        error('Wrong number of value objects for this data bar rule');
      end if;
      
      -- validate the first two VOs
      for i in 1 .. 2 loop
        if not ExcelTypes.isValidCondFmtVOType(rule.cfvoList(i).type) then
          error('Invalid value object type for this data bar rule');
        end if;
      end loop;
      -- validate that the 3rd VO specifies a color
      if rule.cfvoList(3).color is null then
        error('Missing color specification for this data bar rule');
      end if;
      rule.cfvoList(3).color := ExcelTypes.validateColor(rule.cfvoList(3).color);
      
      rule.hideValue := nvl(p_hideValue, false);
    
    when ExcelTypes.CF_TYPE_ICONSET then
      
      rule.type := ExcelTypes.CF_TYPE_MULTISTATE;
      rule.template := ExcelTypes.CF_TEMP_MULTISTATE;
      
      if not ExcelTypes.isValidCondFmtIconSet(p_iconSet) then
        error('Invalid icon set for this conditional formatting rule');
      end if;
      
      rule.iconSet := p_iconSet;
      
      requiredCfvoCount := case 
                           when p_iconSet <= ExcelTypes.CF_ICONSET_3SYMBOLS2 then 2
                           when p_iconSet <= ExcelTypes.CF_ICONSET_4TRAFFICLIGHTS then 3
                           else 4
                           end;
                           
      rule.cfvoList := p_cfvoList;
      
      if rule.cfvoList is null or rule.cfvoList.count = 0 then
        error('No value objects found for this icon set rule');
      elsif rule.cfvoList.count != requiredCfvoCount then
        error('Wrong number of value objects for this icon set rule');
      end if;
      
      rule.hideValue := nvl(p_hideValue, false);
      rule.reverse := nvl(p_reverse, false);
    
    end case;

    return rule;
  end;

  procedure writeDataValidationRule (
    stream  in out nocopy stream_t
  , dvRule  in ExcelTypes.CT_DataValidation
  )
  is
  begin
    stream_write(stream, '<dataValidation type="'||dvRule.type||'"');
    if dvRule.type not in ('list','custom') then
      stream_write(stream, ' operator="'||dvRule.operator||'"');
    end if;
    stream_write(stream, ' allowBlank="'||case when dvRule.allowBlank then '1' else '0' end||'"');
    if dvRule.type = 'list' then
      -- contrary to what its name implies, the showDropDown flag must be 0 to display the dropdown box,
      -- this behaviour is consistent with the XLSB implementation where the flag is named fSuppressCombo
      stream_write(stream, ' showDropDown="'||case when dvRule.showDropDown then '0' else '1' end||'"');
    end if;
    stream_write(stream, ' showInputMessage="'||case when dvRule.showInputMessage then '1' else '0' end||'"');
    stream_write(stream, ' prompt="'||dbms_xmlgen.convert(dvRule.prompt)||'"');
    stream_write(stream, ' promptTitle="'||dbms_xmlgen.convert(dvRule.promptTitle)||'"');
    stream_write(stream, ' showErrorMessage="'||case when dvRule.showErrorMessage then '1' else '0' end||'"');
    stream_write(stream, ' error="'||dbms_xmlgen.convert(dvRule.error)||'"');
    stream_write(stream, ' errorTitle="'||dbms_xmlgen.convert(dvRule.errorTitle)||'"');
    stream_write(stream, ' errorStyle="'||dvRule.errorStyle||'"');
    stream_write(stream, ' sqref="');
    for i in 1 .. dvRule.sqref.ranges.count loop
      if i > 1 then
        stream_write(stream, ' ');
      end if;
      stream_write(stream, dvRule.sqref.ranges(i).value);
    end loop;
    stream_write(stream, '">');
    stream_write(stream, '<formula1>');
    stream_write(stream, dbms_xmlgen.convert(ExcelFmla.parse(dvRule.fmla1, ExcelFmla.FMLATYPE_DATAVAL, dvRule.sqref.lastRangeCellRef, dvRule.refStyle1, dvRule.sqref.boundingAreaFirstCellRef)));
    stream_write(stream, '</formula1>');
    if dvRule.fmla2 is not null then
      stream_write(stream, '<formula2>');
      stream_write(stream, dbms_xmlgen.convert(ExcelFmla.parse(dvRule.fmla2, ExcelFmla.FMLATYPE_DATAVAL, dvRule.sqref.lastRangeCellRef, dvRule.refStyle2, dvRule.sqref.boundingAreaFirstCellRef)));
      stream_write(stream, '</formula2>');
    end if;
    stream_write(stream, '</dataValidation>');
  end;

  procedure writeCondFmtRule (
    stream  in out nocopy stream_t
  , rule    in ExcelTypes.CT_CfRule
  , pos     in pls_integer
  )
  is
  begin
    stream_write(stream, '<conditionalFormatting sqref="');
    for i in 1 .. rule.sqref.ranges.count loop
      if i > 1 then
        stream_write(stream, ' ');
      end if;
      stream_write(stream, rule.sqref.ranges(i).value);
    end loop;
    stream_write(stream, '">');
    
    stream_write(stream, '<cfRule');
    
    -- BEGIN Attributes
    stream_write(stream, ' type="'||ExcelTypes.getCondFmtRuleType(rule.type, rule.template)||'"');
    stream_write(stream, ' priority="'||to_char(pos)||'"');
    
    if rule.dxfId is not null then
      stream_write(stream, ' dxfId="'||to_char(rule.dxfId)||'"');
    end if;
    
    if rule.template in (ExcelTypes.CF_TEMP_EXPR, ExcelTypes.CF_TEMP_CONTAINSTEXT) then
      stream_write(stream, ' operator="'||ExcelTypes.getCondFmtOperator(rule.template, rule.param)||'"');
    end if;
    
    if rule.template = ExcelTypes.CF_TEMP_CONTAINSTEXT then
      stream_write(stream, ' text="'||dbms_xmlgen.convert(rule.strParam)||'"');
    end if;
    
    if rule.type = ExcelTypes.CF_TYPE_EXPRIS 
       and rule.template between ExcelTypes.CF_TEMP_TIMEPERIODTODAY and ExcelTypes.CF_TEMP_TIMEPERIODTHISMONTH
    then
      stream_write(stream, ' timePeriod="'||ExcelTypes.getCondFmtTimePeriod(rule.param)||'"');
    end if;
    
    if rule.template = ExcelTypes.CF_TEMP_BELOWAVERAGE then
      stream_write(stream, ' aboveAverage="0"');
    end if;
    
    if rule.template in (ExcelTypes.CF_TEMP_EQUALBELOWAVERAGE, ExcelTypes.CF_TEMP_EQUALABOVEAVERAGE) then
      stream_write(stream, ' equalAverage="1"');
    end if;
    
    if rule.template in (ExcelTypes.CF_TEMP_BELOWAVERAGE, ExcelTypes.CF_TEMP_ABOVEAVERAGE) 
       and rule.param != 0  
    then
      stream_write(stream, ' stdDev="'||to_char(rule.param)||'"');
    end if;
    
    if rule.template = ExcelTypes.CF_TEMP_FILTER then
      if rule.percent then
        stream_write(stream, ' percent="1"');
      end if;
      if rule.bottom then
        stream_write(stream, ' bottom="1"');
      end if;
      stream_write(stream, ' rank="'||to_char(rule.param)||'"');
    end if;
    
    if rule.stopTrue then
      stream_write(stream, ' stopIfTrue="1"');
    end if;
    -- END Attributes
    
    stream_write(stream, '>');
    
    -- BEGIN Children
    if rule.template = ExcelTypes.CF_TEMP_EXPR then
      stream_write(stream, '<formula>'||dbms_xmlgen.convert(ExcelFmla.parse(rule.fmla1, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle1))||'</formula>');
      if rule.param in (ExcelTypes.CF_OPER_BN, ExcelTypes.CF_OPER_NB) then
        stream_write(stream, '<formula>'||dbms_xmlgen.convert(ExcelFmla.parse(rule.fmla2, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle2))||'</formula>');
      end if;
    
    elsif rule.template = ExcelTypes.CF_TEMP_FMLA 
       or rule.template between ExcelTypes.CF_TEMP_CONTAINSTEXT and ExcelTypes.CF_TEMP_TIMEPERIODTHISMONTH
    then
      stream_write(stream, '<formula>'||dbms_xmlgen.convert(ExcelFmla.parse(rule.fmla1, ExcelFmla.FMLATYPE_CONDFMT, rule.sqref.boundingAreaFirstCellRef, rule.refStyle1))||'</formula>');
      
    elsif rule.template = ExcelTypes.CF_TEMP_GRADIENT then
      stream_write(stream, '<colorScale>');
      
      -- cfvo
      for i in 1 .. rule.cfvoList.count loop
        stream_write(stream, '<cfvo type="'||ExcelTypes.getCondFmtVOType(rule.cfvoList(i).type)||'"');
        if rule.cfvoList(i).type not in (ExcelTypes.CFVO_MIN, ExcelTypes.CFVO_MAX) then
          stream_write(stream, ' val="'||dbms_xmlgen.convert(ExcelFmla.parse(rule.cfvoList(i).value, ExcelFmla.FMLATYPE_CELL, rule.sqref.boundingAreaFirstCellRef, rule.cfvoList(i).refStyle))||'"');
        end if;
        stream_write(stream, '/>');
      end loop;
      -- color
      for i in 1 .. rule.cfvoList.count loop
        stream_write(stream, '<color rgb="'||rule.cfvoList(i).color||'"/>');
      end loop;
      
      stream_write(stream, '</colorScale>');
    
    elsif rule.template = ExcelTypes.CF_TEMP_DATABAR then
      stream_write(stream, '<dataBar'||case when rule.hideValue then ' showValue="0"' end||'>');
      
      -- cfvo
      for i in 1 .. 2 loop
        stream_write(stream, '<cfvo type="'||ExcelTypes.getCondFmtVOType(rule.cfvoList(i).type)||'"');
        if rule.cfvoList(i).type not in (ExcelTypes.CFVO_MIN, ExcelTypes.CFVO_MAX) then
          stream_write(stream, ' val="'||dbms_xmlgen.convert(ExcelFmla.parse(rule.cfvoList(i).value, ExcelFmla.FMLATYPE_CELL, rule.sqref.boundingAreaFirstCellRef, rule.cfvoList(i).refStyle))||'"');
        end if;
        stream_write(stream, '/>');
      end loop;
      -- color
      stream_write(stream, '<color rgb="'||rule.cfvoList(3).color||'"/>');
      
      stream_write(stream, '</dataBar>');
      
    elsif rule.template = ExcelTypes.CF_TEMP_MULTISTATE then
      stream_write(stream, '<iconSet iconSet="'||ExcelTypes.getCondFmtIconSet(rule.iconSet)||'"');
      if rule.hideValue then
        stream_write(stream, ' showValue="0"');
      end if;
      if rule.reverse then
        stream_write(stream, ' reverse="1"');
      end if;
      stream_write(stream, '>');
      
      -- cfvo
      stream_write(stream, '<cfvo type="percent" val="0"/>');
      for i in 1 .. rule.cfvoList.count loop
        stream_write(stream, '<cfvo type="'||ExcelTypes.getCondFmtVOType(rule.cfvoList(i).type)||'"');
        if not nvl(rule.cfvoList(i).gte, true) then
          stream_write(stream, ' gte="0"');
        end if;
        stream_write(stream, ' val="'||dbms_xmlgen.convert(ExcelFmla.parse(rule.cfvoList(i).value, ExcelFmla.FMLATYPE_CELL, rule.sqref.boundingAreaFirstCellRef, rule.cfvoList(i).refStyle))||'"');
        stream_write(stream, '/>');
      end loop;
      
      stream_write(stream, '</iconSet>');
          
    end if;
    -- END Children
    
    stream_write(stream, '</cfRule>');
    
    stream_write(stream, '</conditionalFormatting>');
    
  end;
  
  procedure createWorksheetImpl (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t
  )
  is
    stream          stream_t;
    dataMap         data_map_t;
    t               table_t;
    nrows           integer;
    rowIdx          pls_integer;
    colIdx          pls_integer;
    columnId        pls_integer;
    r               row_t;
    cell            cell_t;
    cell2           floatingCell_t;
    cellSpan        cellSpan_t;
    tableId         pls_integer;
    rId             varchar2(256);
    si              pls_integer;
    
    part            part_t;
    sheet           CT_Sheet;
    
    partitionStart  pls_integer;
    partitionStop   pls_integer;
    
    hasRange        boolean;
    headerXfId      pls_integer;
    rowProps        rowProperties_t;
    
    emptyPartition  exception;
    
    dvRules         ExcelTypes.CT_DataValidations := sd.dvRules;
    cfRules         ExcelTypes.CT_CfRules := sd.cfRules;
    
    bgImage         image_t;
    
  begin
    
    sheet.name := sd.sheetName;
    sheet.tableParts := CT_TableParts();
    sd.sharedFmlaMap.delete;
    
    stream := new_stream(); 
    stream_write(stream, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
      
    if sd.tabColor is not null then
      stream_write(stream, '<sheetPr><tabColor rgb="'||sd.tabColor||'"/></sheetPr>');
    end if;
    
    if sd.hasProps then
      stream_write(stream, 
      '<sheetViews><sheetView workbookViewId="0"' ||
                 case when not sd.showGridLines then ' showGridLines="0"' end ||
                 case when not sd.showRowColHeaders then ' showRowColHeaders="0"' end || '>');
      if sd.activePaneAnchorRef.value is not null then
        stream_write(stream, 
        '<pane xSplit="'||to_char(sd.activePaneAnchorRef.cn - 1)||'" '||
              'ySplit="'||to_char(sd.activePaneAnchorRef.r - 1)||'" '||
              'topLeftCell="'||sd.activePaneAnchorRef.value||'" activePane="bottomLeft" state="frozen"/>');
      end if;
      stream_write(stream, '</sheetView></sheetViews>');
    end if;
    
    -- sheetFormatPr
    if sd.defaultRowHeight is not null then
      stream_write(stream, '<sheetFormatPr defaultRowHeight="'||to_char(sd.defaultRowHeight,'TM9',NLS_PARAM_STRING)||'" customHeight="1"/>');
    end if;
      
    -- columns
    if sd.hasCustomColProps and sd.columnMap.count != 0 then
      stream_write(stream, '<cols>');
      
      columnId := sd.columnMap.first;
      while columnId is not null loop
        stream_write(stream, '<col min="'||to_char(columnId)||'"'||
                                 ' max="'||to_char(columnId)||'"'||
                                 ' width="'||to_char(getColumnWidth(nvl(sd.columnMap(columnId).width, DEFAULT_COL_WIDTH)),'TM9',NLS_PARAM_STRING)||'"'
                                           || case when sd.columnMap(columnId).width is not null then ' customWidth="1"' end ||
                                 case when sd.columnMap(columnId).xfId is not null then
                                 ' style="'||to_char(sd.columnMap(columnId).xfId)||'"'
                                 end || 
                              '/>');
        columnId := sd.columnMap.next(columnId);
      end loop;
      
      stream_write(stream, '</cols>');
    end if;
    
    -- BEGIN sheetData
    stream_write(stream, '<sheetData>');
    
    for tId in 1 .. sd.tableList.count loop
      
      t := sd.tableList(tId);  
      setAnchorRowOffset(sd, t.anchorRef);
      r.id := t.anchorRef.rowOffset - 1;
      cell.r := r.id;
      cell.isTableCell := true;
      
      -- header row
      if t.header.show then
        r.id := r.id + 1;
        -- common cell attributes 
        cell.r := r.id;
                
        if t.rowMap.exists(0) and t.rowMap(0).xfId is not null then
          headerXfId := t.rowMap(0).xfId;
        else
          headerXfId := 0;
        end if;
        
        cell.v.st := ST_STRING;
        
        if sd.streamable then
          writeRowStart(stream, r);
        end if;
        
        for i in 1 .. t.sqlMetadata.columnList.count loop
          if not t.sqlMetadata.columnList(i).excluded then
            columnId := t.sqlMetadata.columnList(i).id;
            cell.cn := t.sqlMetadata.columnList(i).colNum;
            cell.c := t.sqlMetadata.columnList(i).colRef;
            cell.v.varchar2_value := t.columnMap(columnId).name;
            
            cell.xfId := headerXfId;
            -- sheet-level column idx
            colIdx := t.anchorRef.colOffset - 1 + columnId;
            
            -- apply column-specific header style, inheriting from table header style
            if t.columnMap.exists(colIdx) and t.columnMap(colIdx).headerXfId is not null then
              cell.xfId := mergeCellStyle(ctx, cell.xfId, t.columnMap(colIdx).headerXfId);
            end if;
            
            -- inherit from sheet column, sheet or workbook-level style
            if sd.columnMap.exists(colIdx) and sd.columnMap(colIdx).xfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.columnMap(colIdx).xfId, cell.xfId);
            elsif sd.defaultXfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.defaultXfId, cell.xfId);
            end if;
            
            if sd.streamable then
              writeCell(ctx, stream, cell);
            else
              sd.data.rows(cell.r).cells(cell.cn) := cell;
            end if;
            
          end if;
        end loop;
        
        if sd.streamable then
          stream_write(stream, '</row>');
        end if;
        
      end if;
            
      -- prefetch
      nrows := dbms_sql.fetch_rows(t.sqlMetadata.cursorNumber);
      t.isEmpty := (nrows = 0);
      
      partitionStart := t.sqlMetadata.r_num + nrows;
      partitionStop := partitionStart + t.sqlMetadata.partitionSize - 1;
      
      if sd.pageable then
        -- discard empty non-first partition
        if t.isEmpty and t.sqlMetadata.partitionId != 0 then
          closeCursor(t.sqlMetadata);
          sd.done := true;
          dbms_lob.freetemporary(stream.content);
          raise emptyPartition;
        end if;
              
        -- compute pageable sheet name here so that it can be passed to formula context
        if t.sqlMetadata.partitionBySize then
          t.sqlMetadata.partitionId := t.sqlMetadata.partitionId + 1;
          -- t is local, don't forget to write it back to sheet def
          sd.tableList(t.id).sqlMetadata.partitionId := t.sqlMetadata.partitionId;
          sheet.name := replace(sheet.name, '${PNUM}', to_char(t.sqlMetadata.partitionId));
          sheet.name := replace(sheet.name, '${PSTART}', to_char(partitionStart));
          sheet.name := replace(sheet.name, '${PSTOP}', to_char(t.sqlMetadata.r_num));
        end if;
          
        debug('Expanded sheet name='||sheet.name);
      
        ExcelFmla.putSheet(sheet.name);
        ExcelFmla.setCurrentSheet(sheet.name);
      end if;      
      
      -- data rows
      while nrows != 0 loop
        
        r.id := r.id + 1;
        cell.r := r.id;
        
        t.sqlMetadata.r_num := t.sqlMetadata.r_num + 1;
        
        -- retrieve this row properties if set, else inherit from table-level settings
        if t.rowMap.exists(t.sqlMetadata.r_num) then
          rowProps.xfId := nvl(t.rowMap(t.sqlMetadata.r_num).xfId, t.rowProps.xfId);
          rowProps.height := nvl(t.rowMap(t.sqlMetadata.r_num).height, t.rowProps.height);
        else
          rowProps := t.rowProps;
        end if;
        
        if sd.streamable then
          -- use row-level height
          r.props.height := rowProps.height;
          writeRowStart(stream, r);
        else
          -- overwrite sheet-level row height
          sd.data.rows(cell.r).props.height := rowProps.height;
        end if;
        
        -- read current row
        dataMap := getSqlData(t.sqlMetadata);
        
        for i in 1 .. t.sqlMetadata.columnList.count loop
          
          if not t.sqlMetadata.columnList(i).excluded then
          
            cell.v := dataMap(i);
            cell.cn := t.sqlMetadata.columnList(i).colNum;
            cell.c := t.sqlMetadata.columnList(i).colRef;    
            cell.xfId := t.sqlMetadata.columnList(i).xfId;

            -- if original SQL type is ANYDATA, and actual value is numeric or datetime, apply default format
            if t.sqlMetadata.columnList(i).supertype = ST_VARIANT and cell.v.st in (ST_NUMBER, ST_DATETIME) then
              cell.xfId := mergeCellFormat(ctx, cell.xfId, getDefaultFormat(ctx, sd, cell.v.db_type));
            end if;

            -- merge table row-level style
            if rowProps.xfId is not null then
              cell.xfId := mergeCellStyle(ctx, rowProps.xfId, cell.xfId);
            end if;
            
            -- (shared) formula
            if cell.v.st = ST_FORMULA then
              cell.f := t.sqlMetadata.columnList(i).fmla;
              
              if t.sqlMetadata.columnList(i).hyperlink and t.sqlMetadata.columnList(i).linkTokens.count != 0 then
                setLinkTokenValues(cell.f.expr, t.sqlMetadata.columnList(i).linkTokens, dataMap);
              end if;
              
              if cell.f.shared then
                if not sd.sharedFmlaMap.exists(cell.f.sharedIdx) then
                  sd.sharedFmlaMap(cell.f.sharedIdx).columnId := i;
                  sd.sharedFmlaMap(cell.f.sharedIdx).tableId := tId;
                  cell.f.hasRef := true;
                else
                  cell.f.hasRef := false;
                end if;
              end if;
            end if;
            
            if sd.streamable then
              writeCell(ctx, stream, cell);
            else
              sd.data.rows(cell.r).cells(cell.cn) := cell;
            end if;
          
          end if;
          
        end loop;
        
        if sd.streamable then
          stream_write(stream, '</row>');
        end if;

        if t.sqlMetadata.r_num = t.sqlMetadata.maxRows then
          -- force closing cursor
          nrows := 0;
          exit;
        end if;
        
        if cell.r = MAX_ROW_NUMBER then
          if not t.sqlMetadata.partitionBySize then
            -- force closing cursor
            nrows := 0;
          end if;
          exit;
        end if;
        
        exit when t.sqlMetadata.r_num = partitionStop;
        
        -- fetch next row
        nrows := dbms_sql.fetch_rows(t.sqlMetadata.cursorNumber);
      
      end loop;
      
      debug(utl_lms.format_message('end fetch: sheetId=%d tableId=%d rowCount=%d', sd.sheetIndex, tId, t.sqlMetadata.r_num));
      
      if nrows = 0 then
        closeCursor(t.sqlMetadata);
        sd.done := true;
      end if;
      
      hasRange := t.header.show or not t.isEmpty or ( t.isEmpty and t.formatAsTable );
      
      if hasRange then
        t.range := makeRangeImpl(
                     t.sqlMetadata.columnList(t.sqlMetadata.visibleColumnSet.first).colRef
                   , t.anchorRef.rowOffset
                   , t.sqlMetadata.columnList(t.sqlMetadata.visibleColumnSet.last).colRef
                   -- if the table is declared over an empty dataset, extends its range by one row down to make it legal in Excel
                   , cell.r + case when t.formatAsTable and t.isEmpty then 1 else 0 end
                   );
      end if;
           
      if t.formatAsTable then
        tableId := addTableLayout(ctx, t.range, t.header.show, t.header.autoFilter, t.tableStyle, t.columnMap, t.tableName
                                 , t.showFirstColumn, t.showLastColumn, t.showRowStripes, t.showColumnStripes);
        sheet.tableParts.extend;
        sheet.tableParts(sheet.tableParts.last) := tableId;
      end if;
      
      -- column-level data validation and conditional formatting
      if not t.isEmpty then
      
        columnId := t.columnMap.first;
        while columnId is not null loop
          
          -- column data validation
          if t.columnMap(columnId).dvRule.type is not null then
            setSqref(
              sqref     => t.columnMap(columnId).dvRule.sqref
            , cellRange => ExcelTypes.ST_Sqref(
                             makeRangeImpl(
                               t.sqlMetadata.visibleColumnSet(columnId)
                             , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                             , t.sqlMetadata.visibleColumnSet(columnId)
                             , t.range.end_ref.r
                             ).expr
                           )
            );
            dvRules.extend;
            dvRules(dvRules.last) := t.columnMap(columnId).dvRule;
          end if;
          
          -- column conditional formatting
          if t.columnMap(columnId).cfRules is not null and t.columnMap(columnId).cfRules.count != 0 then
            for i in 1 .. t.columnMap(columnId).cfRules.count loop     
              setSqref(
                sqref     => t.columnMap(columnId).cfRules(i).sqref
              , cellRange => ExcelTypes.ST_Sqref(
                               makeRangeImpl(
                                 t.sqlMetadata.visibleColumnSet(columnId)
                               , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                               , t.sqlMetadata.visibleColumnSet(columnId)
                               , t.range.end_ref.r
                               ).expr
                             )
              );
              cfRules.extend;
              cfRules(cfRules.last) := t.columnMap(columnId).cfRules(i);
            end loop;
          end if;
          
          columnId := t.columnMap.next(columnId);
        end loop;
        
        -- table-level conditional formatting
        if t.cfRules is not null and t.cfRules.count != 0 then
          for i in 1 .. t.cfRules.count loop
            setSqref(
              sqref     => t.cfRules(i).sqref
            , cellRange => ExcelTypes.ST_Sqref(
                             makeRangeImpl(
                               t.range.start_ref.c
                             , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                             , t.range.end_ref.c
                             , t.range.end_ref.r
                             ).expr
                           )
            );
            cfRules.extend;
            cfRules(cfRules.last) := t.cfRules(i);        
          end loop;
        end if;
      
      end if;
      
      sd.tableList(tId) := t;

    end loop;
    
    cell.isTableCell := false;
    
    -- resolve floating cells
    for i in 1 .. sd.floatingCells.count loop
      cell2 := sd.floatingCells(i);
      setAnchorRowOffset(sd, cell2.anchorRef);
      setAnchorColOffset(sd, cell2.anchorRef);
      
      cell.r := cell2.anchorRef.rowOffset;
      cell.cn := cell2.anchorRef.colOffset;
      cell.c := base26encode(cell.cn);
      cell.xfId := cell2.xfId;    
      cell.v := cell2.data;
      cell.f := cell2.fmla;
      cell.hyperlink := cell2.hyperlink;
      sd.data.rows(cell.r).id := cell.r;
      sd.data.rows(cell.r).cells(cell.cn) := cell;
    end loop;
    
    -- ranges
    applyRangeStyles(ctx, sd);
    
    -- write in-memory cells
    if sd.data.rows.count != 0 then
      rowIdx := sd.data.rows.first;
      while rowIdx is not null loop

        if sd.data.rows(rowIdx).id is null then
          sd.data.rows(rowIdx).id := rowIdx;
        end if;
        
        writeRowStart(stream, sd.data.rows(rowIdx));
        
        -- cells 
        colIdx := sd.data.rows(rowIdx).cells.first;
        while colIdx is not null loop
          
          cell := sd.data.rows(rowIdx).cells(colIdx);
          
          -- table cell style has already been dealt with earlier
          if not cell.isTableCell then
            
            -- apply number format if needed
            if cell.v.st in (ST_NUMBER, ST_DATETIME) then
              cell.xfId := mergeCellFormat(ctx, cell.xfId, getDefaultFormat(ctx, sd, cell.v.db_type));
            end if;
            -- inherit column-level or sheet-level style
            if sd.columnMap.exists(colIdx) and sd.columnMap(colIdx).xfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.columnMap(colIdx).xfId, cell.xfId);
            elsif sd.defaultXfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.defaultXfId, cell.xfId);
            end if;
            
          end if;
          
          -- inherit row-level style
          if sd.data.rows(rowIdx).props.xfId is not null then
            cell.xfId := mergeCellStyle(ctx, sd.data.rows(rowIdx).props.xfId, cell.xfId);
          end if;
          
          -- master hyperlink style
          if cell.hyperlink then
            cell.xfId := mergeLinkFont(ctx, ctx.workbook.styles.hlinkXfId, cell.xfId);
          end if;
          
          writeCell(ctx, stream, cell);
                  
          colIdx := sd.data.rows(rowIdx).cells.next(colIdx);
            
        end loop;
          
        stream_write(stream, '</row>');
        
        rowIdx := sd.data.rows.next(rowIdx);
                  
      end loop;
      
      sd.done := true;
    
    end if;
    
    stream_write(stream, '</sheetData>');
    -- END sheetData
    
    if sd.tableList.count = 0 then
      sd.done := true;
    end if;
    
    -- if there's only one table, set sheet-level autoFilter accordingly
    if sd.tableList.count = 1 and t.header.show and t.header.autoFilter then
      if not t.formatAsTable then
        
        putNameImpl(
          ctxId     => currentCtxId
        , name      => '_xlnm._FilterDatabase'
        , value     => '''' || sheet.name || '''!' || getRangeExpr(t.range, true)
        , sheetName => sheet.name
        , hidden    => true
        , builtIn   => true
        );
          
        stream_write(stream, '<autoFilter ref="'||getRangeExpr(t.range)||'"/>');
          
      end if;
    end if;
      
    -- merged cells
    if sd.mergedCells.count != 0 then
      stream_write(stream, '<mergeCells count="'||to_char(sd.mergedCells.count)||'">');
      for i in 1 .. sd.mergedCells.count loop
        cellSpan := sd.mergedCells(i);
        setAnchorRowOffset(sd, cellSpan.anchorRef);
        setAnchorColOffset(sd, cellSpan.anchorRef);
        stream_write(stream, '<mergeCell ref="'||makeRangeImpl(cellSpan).expr||'"/>');
      end loop;
      stream_write(stream, '</mergeCells>');
    end if;
      
    -- new sheet
    ctx.workbook.sheets.extend;
    sheet.sheetId := ctx.workbook.sheets.last;

    -- check name validity
    if translate(sheet.name, '_\/*?:[]', '_') != sheet.name 
       or substr(sheet.name, 1, 1) = '''' 
       or substr(sheet.name, -1) = ''''
       or length(sheet.name) > 31 
    then
      error('Invalid sheet name: %s', sheet.name);
    end if;
        
    -- check name uniqueness (case-insensitive)
    if ctx.workbook.sheetMap.exists(upper(sheet.name)) then
      error('Duplicate sheet name: %s', sheet.name);
    end if;
      
    sheet.state := sd.state;
    -- save idx of the first visible sheet
    if ctx.workbook.firstSheet is null and sheet.state = ST_VISIBLE then
      ctx.workbook.firstSheet := sheet.sheetId;
    end if;
        
    sheet.partName := 'xl/worksheets/sheet'||to_char(sheet.sheetId)||'.xml';

    -- new sheet part
    part := new_part(sheet.partName, MT_WORKSHEET);
      
    -- conditional formatting
    for i in 1 .. cfRules.count loop
      writeCondFmtRule(stream, cfRules(i), i);
    end loop;
      
    -- data validations
    if dvRules.count != 0 then
      stream_write(stream, '<dataValidations count="'||to_char(dvRules.count)||'">');
      for i in 1 .. dvRules.count loop
        writeDataValidationRule(stream, dvRules(i));
      end loop;
      stream_write(stream, '</dataValidations>');
    end if;
    
    -- drawings
    if sd.drawing.count != 0 then
      rId := addRelationship(part, RS_DRAWING, createDrawingPart(ctx, sd.drawing));
      stream_write(stream, '<drawing r:id="'||rId||'"/>');
    end if;
    
    -- background image
    if sd.background is not null then
      bgImage := createImagePart(ctx, sd.background);
      rId := addRelationship(part, RS_IMAGE, bgImage.partName);
      stream_write(stream, '<picture r:id="'||rId||'"/>');
    end if;
      
    -- table parts
    if sheet.tableParts.count != 0 then
      stream_write(stream, '<tableParts count="'||to_char(sheet.tableParts.count)||'">');
      for i in 1 .. sheet.tableParts.count loop
        rId := addRelationship(part, RS_TABLE, ctx.workbook.tables(sheet.tableParts(i)).partName);
        stream_write(stream, '<tablePart r:id="'||rId||'"/>');
      end loop;
      stream_write(stream, '</tableParts>');
    end if; 

    stream_write(stream, '</worksheet>');
    stream_flush(stream);
      
    -- set shared formula ranges
    si := sd.sharedFmlaMap.first;
    while si is not null loop
      t := sd.tableList(sd.sharedFmlaMap(si).tableId);
      stream.content := replace(stream.content
                              , '###'||to_char(si)||'###'
                              , makeRangeImpl(t.sqlMetadata.visibleColumnSet(sd.sharedFmlaMap(si).columnId)
                                        , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                                        , t.sqlMetadata.visibleColumnSet(sd.sharedFmlaMap(si).columnId)
                                        , t.range.end_ref.r).expr);
      si := sd.sharedFmlaMap.next(si);
    end loop;
      
    part.content := stream.content;
        
    -- add sheet to workbook
    ctx.workbook.sheets(sheet.sheetId) := sheet;
    ctx.workbook.sheetMap(upper(sheet.name)) := sheet.sheetId;
        
    -- add sheet part to package
    addPart(ctx, part);
    
  exception
    when emptyPartition then
      null;
  end;

  procedure createWorksheetBinImpl (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t
  )
  is
    stream          xutl_xlsb.Stream_T;
    dataMap         data_map_t;
    t               table_t;
    nrows           integer;
    rowIdx          integer := t.anchorRef.rowOffset - 1;
    colIdx          pls_integer;
    columnId        pls_integer;
    r               row_t;
    cell            cell_t;
    cell2           floatingCell_t;
    cellSpan        cellSpan_t;
    tableId         pls_integer;
    rId             varchar2(256);
    si              pls_integer;
    
    part            part_t;
    sheet           CT_Sheet;
    
    partitionStart  pls_integer;
    partitionStop   pls_integer;
    
    hasRange        boolean;
    headerXfId      pls_integer;
    rowProps        rowProperties_t;
    
    emptyPartition  exception;

    dvRules         ExcelTypes.CT_DataValidations := sd.dvRules;
    cfRules         ExcelTypes.CT_CfRules := sd.cfRules;
    
    bgImage         image_t;
    
  begin
    
    sheet.name := sd.sheetName;
    sheet.tableParts := CT_TableParts();
    sd.sharedFmlaMap.delete;
    
    stream := xutl_xlsb.new_stream();
    xutl_xlsb.put_BeginSheet(stream);
      
    if sd.tabColor is not null then
      xutl_xlsb.put_WsProp(stream, sd.tabColor);
    end if;

    if sd.hasProps then   
      xutl_xlsb.put_simple_record(stream, 133);  -- BrtBeginWsViews
      xutl_xlsb.put_BeginWsView(stream, sd.showGridLines, sd.showRowColHeaders);  -- BrtBeginWsView
      if sd.activePaneAnchorRef.value is not null then
        -- BrtPane : 
        xutl_xlsb.put_FrozenPane(stream
                               , numRows => sd.activePaneAnchorRef.r - 1   -- num of frozen rows (ySplit)
                               , numCols => sd.activePaneAnchorRef.cn - 1  -- num of frozen columns  (xSplit)
                               , topRow  => sd.activePaneAnchorRef.r - 1   -- first row of bottom-right pane
                               , leftCol => sd.activePaneAnchorRef.cn - 1  -- first column of bottom-right pane
                               );
      end if;
      xutl_xlsb.put_simple_record(stream, 138);  -- BrtEndWsView
      xutl_xlsb.put_simple_record(stream, 134);  -- BrtEndWsViews
    end if;

    -- sheetFormatPr
    if sd.defaultRowHeight is not null then
      xutl_xlsb.put_WsFmtInfo(stream, sd.defaultRowHeight);
    end if;

    -- columns
    if sd.hasCustomColProps and sd.columnMap.count != 0 then
      xutl_xlsb.put_simple_record(stream, 390);  -- BrtBeginColInfos
      
      columnId := sd.columnMap.first;
      while columnId is not null loop
        xutl_xlsb.put_ColInfo( stream
                             , columnId - 1
                             , colWidth      => getColumnWidth(nvl(sd.columnMap(columnId).width, DEFAULT_COL_WIDTH))
                             , isCustomWidth => ( sd.columnMap(columnId).width is not null )
                             , styleRef      => nvl(sd.columnMap(columnId).xfId, 0)
                             );
        columnId := sd.columnMap.next(columnId);
      end loop;
      
      xutl_xlsb.put_simple_record(stream, 391);  -- BrtEndColInfos
    end if;
      
    -- BEGIN sheetData
    xutl_xlsb.put_simple_record(stream, 145);  -- BrtBeginSheetData
    
    for tId in 1 .. sd.tableList.count loop
      
      t := sd.tableList(tId);
      setAnchorRowOffset(sd, t.anchorRef);
      r.id := t.anchorRef.rowOffset - 1;
      cell.r := r.id;
      cell.isTableCell := true;
      
      -- header row
      if t.header.show then
        r.id := r.id + 1;
        -- common cell attributes 
        cell.r := r.id;
        
        if t.rowMap.exists(0) and t.rowMap(0).xfId is not null then
          headerXfId := t.rowMap(0).xfId;
        else
          headerXfId := 0;
        end if;
        
        cell.v.st := ST_STRING;
        
        if sd.streamable then
          writeRowBin(stream, r, sd.defaultRowHeight);
        end if;
        
        for i in 1 .. t.sqlMetadata.columnList.count loop
          if not t.sqlMetadata.columnList(i).excluded then
            columnId := t.sqlMetadata.columnList(i).id;
            cell.cn := t.sqlMetadata.columnList(i).colNum;
            cell.c := t.sqlMetadata.columnList(i).colRef;
            cell.v.varchar2_value := t.columnMap(columnId).name;

            cell.xfId := headerXfId;
            -- sheet-level column idx
            colIdx := t.anchorRef.colOffset - 1 + columnId;
            
            -- apply column-specific header style, inheriting from table header style
            if t.columnMap.exists(colIdx) and t.columnMap(colIdx).headerXfId is not null then
              cell.xfId := mergeCellStyle(ctx, cell.xfId, t.columnMap(colIdx).headerXfId);
            end if;
            
            -- inherit from sheet column, sheet or workbook-level style            
            if sd.columnMap.exists(colIdx) and sd.columnMap(colIdx).xfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.columnMap(colIdx).xfId, cell.xfId);
            elsif sd.defaultXfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.defaultXfId, cell.xfId);
            end if;
            
            if sd.streamable then
              writeCellBin(ctx, stream, cell);
            else
              sd.data.rows(cell.r).cells(cell.cn) := cell;
            end if;
            
          end if;
        end loop;
      end if;

      -- prefetch
      nrows := dbms_sql.fetch_rows(t.sqlMetadata.cursorNumber);
      t.isEmpty := (nrows = 0);
      
      partitionStart := t.sqlMetadata.r_num + nrows;
      partitionStop := partitionStart + t.sqlMetadata.partitionSize - 1;
      
      if sd.pageable then
        -- discard empty non-first partition
        if t.isEmpty and t.sqlMetadata.partitionId != 0 then
          closeCursor(t.sqlMetadata);
          sd.done := true;
          dbms_lob.freetemporary(stream.content);
          raise emptyPartition;
        end if;
              
        -- compute pageable sheet name here so that it can be passed to formula context
        if t.sqlMetadata.partitionBySize then
          t.sqlMetadata.partitionId := t.sqlMetadata.partitionId + 1;
          -- t is local, don't forget to write it back to sheet def
          sd.tableList(t.id).sqlMetadata.partitionId := t.sqlMetadata.partitionId;
          sheet.name := replace(sheet.name, '${PNUM}', to_char(t.sqlMetadata.partitionId));
          sheet.name := replace(sheet.name, '${PSTART}', to_char(partitionStart));
          sheet.name := replace(sheet.name, '${PSTOP}', to_char(t.sqlMetadata.r_num));
        end if;
          
        debug('Expanded sheet name='||sheet.name);
      
        ExcelFmla.putSheet(sheet.name);
        ExcelFmla.setCurrentSheet(sheet.name);
      end if;
      
      -- data rows
      while nrows != 0 loop
        
        r.id := r.id + 1;
        cell.r := r.id;
        
        t.sqlMetadata.r_num := t.sqlMetadata.r_num + 1;

        -- retrieve this row properties if set, else inherit from table-level settings
        if t.rowMap.exists(t.sqlMetadata.r_num) then
          rowProps.xfId := nvl(t.rowMap(t.sqlMetadata.r_num).xfId, t.rowProps.xfId);
          rowProps.height := nvl(t.rowMap(t.sqlMetadata.r_num).height, t.rowProps.height);
        else
          rowProps := t.rowProps;
        end if;
        
        if sd.streamable then
          -- use row-level height
          r.props.height := rowProps.height;
          writeRowBin(stream, r, sd.defaultRowHeight);
        else
          -- overwrite sheet-level row height
          sd.data.rows(cell.r).props.height := rowProps.height;
        end if;
        
        -- read current row
        dataMap := getSqlData(t.sqlMetadata);
        
        for i in 1 .. t.sqlMetadata.columnList.count loop
          
          if not t.sqlMetadata.columnList(i).excluded then
          
            cell.v := dataMap(i);
            cell.cn := t.sqlMetadata.columnList(i).colNum;
            cell.c := t.sqlMetadata.columnList(i).colRef;
            cell.xfId := t.sqlMetadata.columnList(i).xfId;

            -- if original SQL type is ANYDATA, and actual value is numeric or datetime, apply default format
            if t.sqlMetadata.columnList(i).supertype = ST_VARIANT and cell.v.st in (ST_NUMBER, ST_DATETIME) then
              cell.xfId := mergeCellFormat(ctx, cell.xfId, getDefaultFormat(ctx, sd, cell.v.db_type));
            end if;
            
            -- merge table row-level style
            if rowProps.xfId is not null then
              cell.xfId := mergeCellStyle(ctx, rowProps.xfId, cell.xfId);
            end if;            

            -- (shared) formula
            if cell.v.st = ST_FORMULA then
              cell.f := t.sqlMetadata.columnList(i).fmla;

              if t.sqlMetadata.columnList(i).hyperlink and t.sqlMetadata.columnList(i).linkTokens.count != 0 then
                setLinkTokenValues(cell.f.expr, t.sqlMetadata.columnList(i).linkTokens, dataMap);
              end if;
              
              if cell.f.shared then
                if not sd.sharedFmlaMap.exists(cell.f.sharedIdx) then
                  sd.sharedFmlaMap(cell.f.sharedIdx).columnId := i;
                  sd.sharedFmlaMap(cell.f.sharedIdx).tableId := tId;
                  cell.f.hasRef := true;
                else
                  cell.f.hasRef := false;
                end if;
              end if;
            end if;
            
            if sd.streamable then             
              writeCellBin(ctx, stream, cell);
            else
              sd.data.rows(cell.r).cells(cell.cn) := cell;
            end if;
          
          end if;
          
        end loop;

        if t.sqlMetadata.r_num = t.sqlMetadata.maxRows then
          -- force closing cursor
          nrows := 0;
          exit;
        end if;
        
        if cell.r = MAX_ROW_NUMBER then
          if not t.sqlMetadata.partitionBySize then
            -- force closing cursor
            nrows := 0;
          end if;
          exit;
        end if;
        
        exit when t.sqlMetadata.r_num = partitionStop;
        
        -- fetch next row
        nrows := dbms_sql.fetch_rows(t.sqlMetadata.cursorNumber);
          
      end loop;
      
      debug(utl_lms.format_message('end fetch: sheetId=%d tableId=%d rowCount=%d', sd.sheetIndex, tId, t.sqlMetadata.r_num));

      if nrows = 0 then
        closeCursor(t.sqlMetadata);
        sd.done := true;
      end if;
      
      hasRange := t.header.show or not t.isEmpty or ( t.isEmpty and t.formatAsTable );

      if hasRange then
        t.range := makeRangeImpl(
                     t.sqlMetadata.columnList(t.sqlMetadata.visibleColumnSet.first).colRef
                   , t.anchorRef.rowOffset
                   , t.sqlMetadata.columnList(t.sqlMetadata.visibleColumnSet.last).colRef
                   -- if the table is declared over an empty dataset, extends its range by one row down to make it legal in Excel
                   , cell.r + case when t.formatAsTable and t.isEmpty then 1 else 0 end
                   );
      end if;

      if t.formatAsTable then
        tableId := addTableLayout(ctx, t.range, t.header.show, t.header.autoFilter, t.tableStyle, t.columnMap, t.tableName
                                 , t.showFirstColumn, t.showLastColumn, t.showRowStripes, t.showColumnStripes);
        sheet.tableParts.extend;
        sheet.tableParts(sheet.tableParts.last) := tableId;
      end if;
      
      -- column-level data validation and conditional formatting
      if not t.isEmpty then
        
        columnId := t.columnMap.first;
        while columnId is not null loop
            
          -- column data validation
          if t.columnMap(columnId).dvRule.type is not null then
            setSqref(
              sqref     => t.columnMap(columnId).dvRule.sqref
            , cellRange => ExcelTypes.ST_Sqref(
                             makeRangeImpl(
                               t.sqlMetadata.visibleColumnSet(columnId)
                             , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                             , t.sqlMetadata.visibleColumnSet(columnId)
                             , t.range.end_ref.r
                             ).expr
                           )
            );
            dvRules.extend;
            dvRules(dvRules.last) := t.columnMap(columnId).dvRule;
          end if;
            
          -- column conditional formatting
          if t.columnMap(columnId).cfRules is not null and t.columnMap(columnId).cfRules.count != 0 then
            for i in 1 .. t.columnMap(columnId).cfRules.count loop     
              setSqref(
                sqref     => t.columnMap(columnId).cfRules(i).sqref
              , cellRange => ExcelTypes.ST_Sqref(
                               makeRangeImpl(
                                 t.sqlMetadata.visibleColumnSet(columnId)
                               , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                               , t.sqlMetadata.visibleColumnSet(columnId)
                               , t.range.end_ref.r
                               ).expr
                             )
              );
              cfRules.extend;
              cfRules(cfRules.last) := t.columnMap(columnId).cfRules(i);
            end loop;
          end if;
            
          columnId := t.columnMap.next(columnId);
        end loop;
          
        -- table-level conditional formatting
        if t.cfRules is not null and t.cfRules.count != 0 then
          for i in 1 .. t.cfRules.count loop
            setSqref(
              sqref     => t.cfRules(i).sqref
            , cellRange => ExcelTypes.ST_Sqref(
                             makeRangeImpl(
                               t.range.start_ref.c
                             , t.range.start_ref.r + case when t.header.show then 1 else 0 end
                             , t.range.end_ref.c
                             , t.range.end_ref.r
                             ).expr
                           )
            );
            cfRules.extend;
            cfRules(cfRules.last) := t.cfRules(i);        
          end loop;
        end if;
      
      end if;
      
      sd.tableList(tId) := t;
      
    end loop;
    
    cell.isTableCell := false;

    -- resolve floating cells
    for i in 1 .. sd.floatingCells.count loop
      cell2 := sd.floatingCells(i);
      setAnchorRowOffset(sd, cell2.anchorRef);
      setAnchorColOffset(sd, cell2.anchorRef);
      
      cell.r := cell2.anchorRef.rowOffset;
      cell.cn := cell2.anchorRef.colOffset;
      cell.c := base26encode(cell.cn);
      cell.xfId := cell2.xfId;    
      cell.v := cell2.data;
      cell.f := cell2.fmla;
      cell.hyperlink := cell2.hyperlink;
      sd.data.rows(cell.r).id := cell.r;
      sd.data.rows(cell.r).cells(cell.cn) := cell;
    end loop;

    -- ranges
    applyRangeStyles(ctx, sd);

    -- write in-memory cells
    if sd.data.rows.count != 0 then
      rowIdx := sd.data.rows.first;
      while rowIdx is not null loop

        if sd.data.rows(rowIdx).id is null then
          sd.data.rows(rowIdx).id := rowIdx;
        end if;
      
        writeRowBin(stream, sd.data.rows(rowIdx), sd.defaultRowHeight);
          
        -- cells 
        colIdx := sd.data.rows(rowIdx).cells.first;
        while colIdx is not null loop
            
          cell := sd.data.rows(rowIdx).cells(colIdx);
          
          -- table cell style has already been dealt with earlier
          if not cell.isTableCell then
            
            -- apply number format if needed
            if cell.v.st in (ST_NUMBER, ST_DATETIME) then
              cell.xfId := mergeCellFormat(ctx, cell.xfId, getDefaultFormat(ctx, sd, cell.v.db_type));
            end if;
            -- inherit column-level or sheet-level style
            if sd.columnMap.exists(colIdx) and sd.columnMap(colIdx).xfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.columnMap(colIdx).xfId, cell.xfId);
            elsif sd.defaultXfId is not null then
              cell.xfId := mergeCellStyle(ctx, sd.defaultXfId, cell.xfId);
            end if; 
            
          end if;
          
          -- inherit row-level style
          if sd.data.rows(rowIdx).props.xfId is not null then
            cell.xfId := mergeCellStyle(ctx, sd.data.rows(rowIdx).props.xfId, cell.xfId);
          end if;

          -- master hyperlink style
          if cell.hyperlink then
            cell.xfId := mergeLinkFont(ctx, ctx.workbook.styles.hlinkXfId, cell.xfId);
          end if;

          writeCellBin(ctx, stream, cell);
                    
          colIdx := sd.data.rows(rowIdx).cells.next(colIdx);
              
        end loop;
          
        rowIdx := sd.data.rows.next(rowIdx);
                    
      end loop;
      
      sd.done := true;
      
    end if;
      
    xutl_xlsb.put_simple_record(stream, 146);  -- BrtEndSheetData
    -- END sheetData
    
    if sd.tableList.count = 0 then
      sd.done := true;
    end if;
        
    -- if there's only one table, set sheet-level autoFilter accordingly
    if sd.tableList.count = 1 and t.header.show and t.header.autoFilter then
      if not t.formatAsTable then
          
        putNameImpl(
          ctxId     => currentCtxId
        , name      => '_FilterDatabase'
        , value     => '''' || sheet.name || '''!' || getRangeExpr(t.range, true)
        , sheetName => sheet.name
        , hidden    => true
        , builtIn   => true
        );

        xutl_xlsb.put_BeginAFilter(
          stream
        , firstRow    => t.range.start_ref.r - 1
        , firstCol    => t.range.start_ref.cn - 1
        , lastRow     => t.range.end_ref.r - 1
        , lastCol     => t.range.end_ref.cn - 1
        );
        xutl_xlsb.put_simple_record(stream, 162);  -- BrtEndAFilter
        
      end if;
    end if;

    -- merged cells
    if sd.mergedCells.count != 0 then
      xutl_xlsb.put_simple_record(stream, 177, int2raw(sd.mergedCells.count)); -- BrtBeginMergeCells
      for i in 1 .. sd.mergedCells.count loop
        cellSpan := sd.mergedCells(i);
        setAnchorRowOffset(sd, cellSpan.anchorRef);
        setAnchorColOffset(sd, cellSpan.anchorRef);
        xutl_xlsb.put_MergeCell(
          stream
        , rwFirst  => cellSpan.anchorRef.rowOffset - 1
        , rwLast   => ( cellSpan.anchorRef.rowOffset + cellSpan.rowSpan - 1 ) - 1
        , colFirst => cellSpan.anchorRef.colOffset - 1
        , colLast  => ( cellSpan.anchorRef.colOffset + cellSpan.colSpan - 1 ) - 1
        );
      end loop;
      xutl_xlsb.put_simple_record(stream, 178); -- BrtEndMergeCells
    end if;
           
    -- new sheet
    ctx.workbook.sheets.extend;
    sheet.sheetId := ctx.workbook.sheets.last;

    -- check name validity
    if translate(sheet.name, '_\/*?:[]', '_') != sheet.name 
       or substr(sheet.name, 1, 1) = '''' 
       or substr(sheet.name, -1) = ''''
       or length(sheet.name) > 31 
    then
      error('Invalid sheet name: %s', sheet.name);
    end if;
          
    -- check name uniqueness (case-insensitive)
    if ctx.workbook.sheetMap.exists(upper(sheet.name)) then
      error('Duplicate sheet name: %s', sheet.name);
    end if;

    sheet.state := sd.state;
    -- save idx of the first visible sheet
    if ctx.workbook.firstSheet is null and sheet.state = ST_VISIBLE then
      ctx.workbook.firstSheet := sheet.sheetId;
    end if;
        
    sheet.partName := 'xl/worksheets/sheet'||to_char(sheet.sheetId)||'.bin';

    -- new sheet part
    part := new_part(sheet.partName, MT_WORKSHEET_BIN);

    -- conditional formatting
    if cfRules.count != 0 then
      xutl_xlsb.put_CondFmts(stream, cfRules);
    end if;
      
    -- data validations
    if dvRules.count != 0 then
      xutl_xlsb.put_DVals(stream, dvRules);
    end if;

    -- drawings
    if sd.drawing.count != 0 then
      rId := addRelationship(part, RS_DRAWING, createDrawingPart(ctx, sd.drawing));
      xutl_xlsb.put_Drawing(stream, rId);
    end if;
    
    -- background image
    if sd.background is not null then
      bgImage := createImagePart(ctx, sd.background);
      rId := addRelationship(part, RS_IMAGE, bgImage.partName);
      xutl_xlsb.put_BkHim(stream, rId);
    end if;
      
    -- table parts
    if sheet.tableParts.count != 0 then
      xutl_xlsb.put_simple_record(stream, 660, int2raw(sheet.tableParts.count)); -- BrtBeginListParts
      for i in 1 .. sheet.tableParts.count loop
        rId := addRelationship(part, RS_TABLE, ctx.workbook.tables(sheet.tableParts(i)).partName);
        xutl_xlsb.put_ListPart(stream, rId);  -- BrtListPart
      end loop;
      xutl_xlsb.put_simple_record(stream, 662);  -- BrtEndListParts
    end if;
        
    xutl_xlsb.put_simple_record(stream, 130);  -- BrtEndSheet
    xutl_xlsb.flush_stream(stream);

    -- set shared formula ranges      
    si := sd.sharedFmlaMap.first;
    while si is not null loop
      t := sd.tableList(sd.sharedFmlaMap(si).tableId);
        
      xutl_xlsb.put_ShrFmlaRfX(
        stream   => stream
      , si       => si
      , firstRow => t.range.start_ref.r + case when t.header.show then 1 else 0 end - 1
      , firstCol => t.sqlMetadata.columnList(sd.sharedFmlaMap(si).columnId).colNum - 1
      , lastRow  => t.range.end_ref.r - 1
      , lastCol  => t.sqlMetadata.columnList(sd.sharedFmlaMap(si).columnId).colNum - 1
      );
        
      si := sd.sharedFmlaMap.next(si);
    end loop;
      

    part.contentBin := stream.content;
    part.isBinary := true;
        
    -- add sheet to workbook
    ctx.workbook.sheets(sheet.sheetId) := sheet;
    ctx.workbook.sheetMap(upper(sheet.name)) := sheet.sheetId;
        
    -- add sheet part to package
    addPart(ctx, part);
    
  exception
    when emptyPartition then
      null;      
  end;
  
  procedure prepareTable (
    ctx  in out nocopy context_t
  , sd   in out nocopy sheet_definition_t
  , i    in pls_integer
  ) 
  is
    defaultFmt        varchar2(128);
    DEFAULT_XF        CT_Xf;
    cellXf            CT_Xf;
    columnId          pls_integer;
    sheetColumnId     pls_integer;
    tableColumn       table_column_t;
    hasTableColProps  boolean;
  begin
    
    setAnchorColOffset(sd, sd.tableList(i).anchorRef);
          
    prepareCursor(sd.tableList(i).sqlMetadata, sd.tableList(i).anchorRef.colOffset);
        
    -- set column-level information
    for j in 1 .. sd.tableList(i).sqlMetadata.columnList.count loop
          
      if not sd.tableList(i).sqlMetadata.columnList(j).excluded then
            
        columnId := sd.tableList(i).sqlMetadata.columnList(j).id; -- visible column ID
        cellXf := DEFAULT_XF;
            
        tableColumn := null;
        hasTableColProps := sd.tableList(i).columnMap.exists(columnId);
        if hasTableColProps then
          tableColumn := sd.tableList(i).columnMap(columnId);
        end if;
          
        if tableColumn.name is null then
          sd.tableList(i).columnMap(columnId).name := sd.tableList(i).sqlMetadata.columnList(j).name;         
        end if;
        
        -- table-level column style
        if tableColumn.xfId is not null then
          cellXf := getCellXf(ctx, tableColumn.xfId);
        end if;
        -- inherit from sheet column style, or sheet style, if defined
        sheetColumnId := columnId + sd.tableList(i).anchorRef.colOffset - 1;
        if sd.columnMap.exists(sheetColumnId) and sd.columnMap(sheetColumnId).xfId is not null then
          mergeCellStyleImpl(ctx, getCellXf(ctx, sd.columnMap(sheetColumnId).xfId), cellXf);
        elsif sd.defaultXfId is not null then
          mergeCellStyleImpl(ctx, getCellXf(ctx, sd.defaultXfId), cellXf);
        end if;
          
        defaultFmt := getDefaultFormat(ctx, sd, sd.tableList(i).sqlMetadata.columnList(j).type);
        -- if no numFmt defined, apply default
        if cellXf.numFmtId = 0 and defaultFmt is not null then
          cellXf.numFmtId := putNumfmt(ctx.workbook.styles, defaultFmt);
        end if;
            
        -- if defined on this column, apply hyperlink master style and font
        if sd.tableList(i).sqlMetadata.columnList(j).hyperlink then
          cellXf.xfId := ctx.workbook.styles.hlinkXfId;
          cellXf.fontId := ctx.workbook.styles.cellStyleXfs(cellXf.xfId).fontId;
          prepareHyperlink(sd.tableList(i).sqlMetadata, j);
        else
          cellXf.xfId := 0; -- Normal style
        end if;
            
        setCellXfContent(cellXf);
        sd.tableList(i).sqlMetadata.columnList(j).xfId := putCellXf(ctx.workbook.styles, cellXf);
          
      end if;
          
    end loop;
    
    -- recursively process dependent tables
    for j in 1 .. sd.tableForest.t(i).children.count loop
      prepareTable(ctx, sd, sd.tableForest.t(i).children(j));
    end loop;
    
  end;

  procedure createWorksheet (
    ctx  in out nocopy context_t
  , sh   in sheetHandle
  )
  is
    sd   sheet_definition_t;
    idx  pls_integer;
  begin
    
    sd := ctx.sheetDefinitionMap(sh);
    -- apply global style to sheet
    sd.defaultXfId := mergeCellStyle(ctx, ctx.defaultXfId, sd.defaultXfId);
    
    -- apply sheet-level styles to lower levels
    if sd.defaultXfId is not null then
      -- row styles
      idx := sd.data.rows.first;
      while idx is not null loop
        if sd.data.rows(idx).props.xfId is not null then
          sd.data.rows(idx).props.xfId := mergeCellStyle(ctx, sd.defaultXfId, sd.data.rows(idx).props.xfId);
        end if;
        idx := sd.data.rows.next(idx);
      end loop;
      -- column styles
      idx := sd.columnMap.first;
      while idx is not null loop
        sd.columnMap(idx).xfId := mergeCellStyle(ctx, sd.defaultXfId, sd.columnMap(idx).xfId);
        idx := sd.columnMap.next(idx);
      end loop;
    end if;
    
    sd.tableForest := getTableForest(sd.tableList);
    
    -- prepare root tables
    for i in 1 .. sd.tableForest.roots.count loop
      prepareTable(ctx, sd, sd.tableForest.roots(i));
    end loop;
    
    sd.streamable := ( sd.tableList.count = 1 
                   and sd.data.rows.count = 0 
                   and sd.floatingCells.count = 0
                   and sd.cellRanges.count = 0 );
    sd.done := false;
    
    -- layout check
    if sd.pageable and not sd.streamable then
      error('Cannot paginate data in a multitable or mixed-content worksheet');
    end if;
    
    -- do not set the current sheet name in formula context when pagination is enabled
    if not sd.pageable then
      ExcelFmla.setCurrentSheet(sd.sheetName);
    end if;
    
    -- for a pageable sheet, append a default partition number placeholder to the sheet name if none found
    if sd.pageable and not regexp_like(sd.sheetName, '\$\{P(NUM|START|STOP)\}') then
      sd.sheetName := sd.sheetName || '_${PNUM}';
    end if;
    
    while not sd.done loop
      case ctx.fileType
      when FILE_XLSX then
        createWorksheetImpl(ctx, sd);
      when FILE_XLSB then
        createWorksheetBinImpl(ctx, sd);
      end case;
    end loop;
    
    ctx.sheetDefinitionMap(sh) := sd;

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
                         case when tab.showFirstColumn then ' showFirstColumn="1"' end ||
                         case when tab.showLastColumn then ' showLastColumn="1"' end ||
                         case when tab.showRowStripes then ' showRowStripes="1"' end ||
                         case when tab.showColumnStripes then ' showColumnStripes="1"' end ||    
                         '/>');
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
    
    xutl_xlsb.put_TableStyleClient(  -- BrtTableStyleClient
      stream
    , tab.styleName
    , tab.showFirstColumn
    , tab.showLastColumn
    , tab.showRowStripes
    , tab.showColumnStripes
    );  
    
    xutl_xlsb.put_simple_record(stream, 344);  -- BrtEndList
    xutl_xlsb.flush_stream(stream);
    addPart(ctx, tab.partName, MT_TABLE_BIN, stream.content);
  end;
  
  procedure createWorkbook (
    ctx   in out nocopy context_t
  )
  is
    stream  stream_t;
    part    part_t := new_part('xl/workbook.xml', MT_WORKBOOK);
  begin
    ctx.workbook.partName := part.name;
    stream := new_stream();
    
    stream_write(stream, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    stream_write(stream, '<fileVersion appName="xl" lastEdited="2" lowestEdited="2"/>');
    
    stream_write(stream, '<bookViews>');
    stream_write(stream, '<workbookView firstSheet="'||to_char(ctx.workbook.firstSheet - 1)||'" activeTab="'||to_char(ctx.workbook.firstSheet - 1)||'"/>');
    stream_write(stream, '</bookViews>');
    
    stream_write(stream, '<sheets>');
    
    for i in 1 .. ctx.workbook.sheets.count loop
      -- add sheet relationships
      ctx.workbook.sheets(i).rId := addRelationship(part, RS_WORKSHEET, ctx.workbook.sheets(i).partName);
      stream_write(stream, '<sheet name="'||dbms_xmlgen.convert(ctx.workbook.sheets(i).name) ||
                               '" sheetId="'||ctx.workbook.sheets(i).sheetId || '"' ||
                               case when ctx.workbook.sheets(i).state != ST_VISIBLE 
                                    then ' state="' ||
                                         case ctx.workbook.sheets(i).state
                                         when ST_HIDDEN then 'hidden'
                                         when ST_VERYHIDDEN then 'veryHidden'
                                         end || '"' 
                               end ||
                               ' r:id="'||ctx.workbook.sheets(i).rId||'"/>');
    end loop;
    
    stream_write(stream, '</sheets>');
    
    if ctx.names.count != 0 then
      stream_write(stream, '<definedNames>');
      for i in 1 .. ctx.names.count loop
        
        if not ctx.names(i).futureFunction then
          
          stream_write(stream, '<definedName name="'||ctx.names(i).name||'"');
          if ctx.names(i).comment is not null then
            stream_write(stream, ' comment="'||dbms_xmlgen.convert(ctx.names(i).comment)||'"');
          end if;
          if ctx.names(i).scope is not null then
            stream_write(stream, ' localSheetId="'||to_char(ctx.workbook.sheetMap(upper(ctx.names(i).scope)) - 1)||'"');
          end if;
          if ctx.names(i).hidden then
            stream_write(stream, ' hidden="1"');
          end if;
          stream_write(stream, '>');
          
          -- set current sheet to resolve unscoped cell references in the formula,
          -- an error will be raised during parsing if an unscoped cell reference is found in a workbook-level name
          ExcelFmla.setCurrentSheet(ctx.names(i).scope);
          stream_write(stream
                     , dbms_xmlgen.convert(
                         ExcelFmla.parse(
                           p_expr     => ctx.names(i).formula
                         , p_type     => ExcelFmla.FMLATYPE_NAME
                         , p_cellRef  => ctx.names(i).cellRef
                         , p_refStyle => ctx.names(i).refStyle
                         ) ) );
          
          stream_write(stream, '</definedName>');
        
        end if;
        
      end loop;
            
      stream_write(stream, '</definedNames>');
      
    end if;
    
    -- set calculation engine version to max value
    stream_write(stream, '<calcPr calcId="999999"' || case when ctx.workbook.refStyle = ExcelFmla.REF_R1C1 then ' refMode="R1C1"' end || '/>');
    
    stream_write(stream, '</workbook>');
    stream_flush(stream);
    
    part.content := stream.content;
    debug(xmltype(part.content).getstringval(1,2));
    
    addPart(ctx, part);
    
    createSharedStrings(ctx);
    addRelationship(ctx, part.name, RS_SHAREDSTRINGS, 'xl/sharedStrings.xml');

    createStylesheet(ctx, ctx.workbook.styles);
    addRelationship(ctx, part.name, RS_STYLES, 'xl/styles.xml');
    
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
    part    part_t := new_part('xl/workbook.bin', null);
  begin
    ctx.workbook.partName := part.name;
    xutl_xlsb.put_simple_record(stream, 131); -- BrtBeginBook
    xutl_xlsb.put_defaultBookViews(stream, ctx.workbook.firstSheet - 1);
    xutl_xlsb.put_simple_record(stream, 143); -- BrtBeginBundleShs
    xutl_xlsb.resetSheetCache;
    
    for i in 1 .. ctx.workbook.sheets.count loop
      -- add sheet relationships
      ctx.workbook.sheets(i).rId := addRelationship(part, RS_WORKSHEET, ctx.workbook.sheets(i).partName);
      xutl_xlsb.put_BundleSh(stream, ctx.workbook.sheets(i).sheetId, ctx.workbook.sheets(i).rId, ctx.workbook.sheets(i).name, ctx.workbook.sheets(i).state);
    end loop;
    
    xutl_xlsb.put_simple_record(stream, 144); -- BrtEndBundleShs
      
    ctx.names := ExcelFmla.getNames(); -- sync names from formula context
    xutl_xlsb.put_Names(stream, ctx.names);
    
    xutl_xlsb.put_CalcProp(stream, 999999, ctx.workbook.refStyle); -- BrtCalcProp
    
    xutl_xlsb.put_simple_record(stream, 132);  -- BrtEndBook
    xutl_xlsb.flush_stream(stream);
    part.contentBin := stream.content;
    part.isBinary := true;
    addPart(ctx, part);
    
    createSharedStringsBin(ctx);
    addRelationship(ctx, part.name, RS_SHAREDSTRINGS, 'xl/sharedStrings.bin');
    
    createStylesheetBin(ctx, ctx.workbook.styles);
    addRelationship(ctx, part.name, RS_STYLES, 'xl/styles.bin');
    
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
    stream       stream_t := new_stream();
    contentType  varchar2(256);
  begin    
    stream_write(stream, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
    -- default extensions
    contentType := ctx.extensions.first;
    while contentType is not null loop
      stream_write(stream, '<Default Extension="'||ctx.extensions(contentType)||'" ContentType="'||contentType||'"/>');
      contentType := ctx.extensions.next(contentType);
    end loop;
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
    
    if dbms_lob.istemporary(binaryContent) = 1 then
      dbms_lob.freetemporary(binaryContent);
    end if;
    
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
        if dbms_lob.istemporary(pck.parts(i).contentBin) = 1 then
          dbms_lob.freetemporary(pck.parts(i).contentBin);
        end if;
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
    ctx.names := ExcelTypes.CT_DefinedNames();
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
    if p_ctxId = currentCtxId then
      currentCtx := null;
      currentCtxId := -1;
    end if;
  end;

  function putTableImpl (
    sd             in out nocopy sheet_definition_t
  , p_query        in clob
  , p_rc           in sys_refcursor
  , p_paginate     in boolean default false
  , p_pageSize     in pls_integer default null
  , p_anchorRef    in anchorRef_t default null
  , p_maxRows      in integer default null
  , p_excludeCols  in varchar2 default null
  )
  return tableHandle
  is
    t         table_t;
    local_rc  sys_refcursor := p_rc;
  begin
    t.formatAsTable := false;
    
    if p_paginate then
      t.sqlMetadata.partitionBySize := true;
      t.sqlMetadata.partitionSize := nvl(p_pageSize, MAX_ROW_NUMBER);
      sd.pageable := true;
    end if;    
    
    if p_query is not null then
      t.sqlMetadata.queryString := p_query;
      t.sqlMetadata.bindVariables := bind_variable_list_t();
    else
      t.sqlMetadata.cursorNumber := dbms_sql.to_cursor_number(local_rc);
    end if;
    
    t.anchorRef := p_anchorRef;
    if t.anchorRef.rowOffset is null then
      t.anchorRef.rowOffset := 1;
    end if;
    if t.anchorRef.colOffset is null then
      t.anchorRef.colOffset := 1;
    end if;
    
    t.sqlMetadata.virtualColumns := virtualColumnList_t();
    t.sqlMetadata.excludeSet := parseValueList(p_excludeCols);
    t.sqlMetadata.maxRows := p_maxRows;
    
    sd.tableList.extend;
    t.id := sd.tableList.last;
    sd.tableList(t.id) := t;
        
    return t.id;
  end;

  function addSheetImpl (
    ctx           in out nocopy context_t
  , p_sheetName   in varchar2
  , p_tabColor    in varchar2 default null
  , p_sheetIndex  in pls_integer default null
  , p_state       in pls_integer default 0
  )
  return sheetHandle
  is
    sd  sheet_definition_t;
  begin    
    if p_sheetIndex is not null then
      if not ctx.sheetDefinitionMap.exists(p_sheetIndex) then
        sd.sheetIndex := p_sheetIndex;
      else
        error('Duplicate sheet index: %d', p_sheetIndex);
      end if;
    else
      sd.sheetIndex := nvl(ctx.sheetDefinitionMap.last, 0) + 1;
    end if;
    
    if ctx.sheetIndexMap.exists(upper(p_sheetName)) then
       error('Duplicate sheet name: %s', p_sheetName);
    end if;
    
    sd.sheetName := p_sheetName;
    sd.tabColor := ExcelTypes.validateColor(p_tabColor);
    if p_state not in (ST_VISIBLE, ST_HIDDEN, ST_VERYHIDDEN) then
      error('Invalid sheet visibility state: %d', p_state);
    end if;
    sd.state := nvl(p_state, ST_VISIBLE);
    
    sd.mergedCells := cellSpanList_t();
    sd.tableList := tableList_t();
    sd.data.hasCells := false;
    sd.floatingCells := floatingCellList_t();
    sd.cellRanges := cellRangeList_t();
    sd.sharedFmlaSeq := 0;
    sd.dvRules := ExcelTypes.CT_DataValidations();
    sd.cfRules := ExcelTypes.CT_CfRules();
    sd.drawing := CT_Drawing();
    
    ctx.sheetDefinitionMap(sd.sheetIndex) := sd;
    ctx.sheetIndexMap(upper(sd.sheetName)) := sd.sheetIndex;
    
    return sd.sheetIndex;
    
  end;

  function addSheet (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_tabColor    in varchar2 default null
  , p_sheetIndex  in pls_integer default null
  , p_state       in pls_integer default null
  )
  return sheetHandle
  is
  begin
    loadContext(p_ctxId);
    return addSheetImpl(
             currentCtx
           , p_sheetName
           , p_tabColor
           , p_sheetIndex
           , p_state
           );
  end;
  
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
  )
  is
    sheetId  sheetHandle;
  begin
    sheetId := addSheetFromQuery(p_ctxId, p_sheetName, p_query, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_maxRows, p_state, p_excludeCols);
  end;

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
  )
  is
  begin
    addSheetFromQuery(p_ctxId, p_sheetName, to_clob(p_query), p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_maxRows, p_state, p_excludeCols);
  end;

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
  return sheetHandle
  is
    sheetId  sheetHandle;
    tableId  tableHandle;
  begin
    loadContext(p_ctxId);
    if p_query is null or dbms_lob.getlength(p_query) = 0 then
      error('Query string argument is null or empty');
    else
      sheetId := addSheetImpl(currentCtx, p_sheetName, p_tabColor, p_sheetIndex, p_state);
    end if;
    
    tableId := putTableImpl(currentCtx.sheetDefinitionMap(sheetId), p_query, null, p_paginate, p_pageSize, null, p_maxRows, p_excludeCols);
    
    return sheetId;
  end;

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
  return sheetHandle
  is
  begin
    return addSheetFromQuery(p_ctxId, p_sheetName, to_clob(p_query), p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_maxRows, p_state, p_excludeCols);
  end;

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
  )
  is
    sheetId  sheetHandle;
  begin
    sheetId := addSheetFromCursor(p_ctxId, p_sheetName, p_rc, p_tabColor, p_paginate, p_pageSize, p_sheetIndex, p_maxRows, p_state, p_excludeCols);
  end;

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
  return sheetHandle
  is
    sheetId  sheetHandle;
    tableId  tableHandle;
  begin
    loadContext(p_ctxId);
    if p_rc is null then
      error('Ref cursor argument cannot be null');
    else
      sheetId := addSheetImpl(currentCtx, p_sheetName, p_tabColor, p_sheetIndex, p_state);
    end if;
    tableId := putTableImpl(currentCtx.sheetDefinitionMap(sheetId), null, p_rc, p_paginate, p_pageSize, null, p_maxRows, p_excludeCols);
    return sheetId;
  end;

  procedure assertTableExists (
    p_sheetId  in sheetHandle
  , p_tableId  in tableHandle   
  )
  is
  begin
    if not currentCtx.sheetDefinitionMap(p_sheetId).tableList.exists(p_tableId) then
      error('Undefined table handle (id=%d)', p_tableId);
    end if;
  end;

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
  return tableHandle
  is
  begin
    return addTable(p_ctxId, p_sheetId, to_clob(p_query), p_paginate, p_pageSize, p_anchorRowOffset, p_anchorColOffset, p_anchorTableId, p_anchorPosition, p_maxRows, p_excludeCols);
  end;

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
  return tableHandle
  is
    tableId    tableHandle;
    anchorRef  anchorRef_t;
  begin
    loadContext(p_ctxId);

    if p_query is null or dbms_lob.getlength(p_query) = 0 then
      error('Query string argument is null or empty');
    end if;

    if p_anchorTableId is not null then
      assertTableExists(p_sheetId, p_anchorTableId);
    else
      assertPositive(p_anchorRowOffset, 'The table anchor row offset must be a positive integer.');
      assertPositive(p_anchorColOffset, 'The table anchor column offset must be a positive integer.');
    end if;

    anchorRef.rowOffset := p_anchorRowOffset;
    anchorRef.colOffset := p_anchorColOffset;
    anchorRef.tableId := p_anchorTableId;
    anchorRef.anchorPosition := p_anchorPosition;

    tableId := putTableImpl(currentCtx.sheetDefinitionMap(p_sheetId), p_query, null, p_paginate, p_pageSize, anchorRef, p_maxRows, p_excludeCols);
    return tableId;
  end;

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
  return tableHandle
  is
    tableId    tableHandle;
    anchorRef  anchorRef_t;
  begin
    loadContext(p_ctxId);

    if p_anchorTableId is not null then
      assertTableExists(p_sheetId, p_anchorTableId);
    else
      assertPositive(p_anchorRowOffset, 'The table anchor row offset must be a positive integer.');
      assertPositive(p_anchorColOffset, 'The table anchor column offset must be a positive integer.');
    end if;
    
    anchorRef.rowOffset := p_anchorRowOffset;
    anchorRef.colOffset := p_anchorColOffset;
    anchorRef.tableId := p_anchorTableId;
    anchorRef.anchorPosition := p_anchorPosition;
    
    tableId := putTableImpl(currentCtx.sheetDefinitionMap(p_sheetId), null, p_rc, p_paginate, p_pageSize, anchorRef, p_maxRows, p_excludeCols);
    return tableId;
  end;

  procedure addTableColumnImpl (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_tableId         in tableHandle
  , p_name            in varchar2
  , p_value           in varchar2
  , p_columnId        in pls_integer
  , p_after           in boolean
  , p_refStyle        in pls_integer default null
  , p_hyperlink       in boolean default false
  )
  is
    t      table_t;
    vc     virtualColumn_t;
  begin
    loadContext(p_ctxId);
    t := currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId);
    
    vc.col.name := p_name;
    vc.col.virtual := true;
    
    if p_value is not null then
    
      vc.col.type := dbms_sql.VARCHAR2_TYPE;
      vc.col.supertype := ST_FORMULA;
      vc.col.fmla.expr := p_value;
      vc.col.fmla.shared := true;
      vc.col.fmla.refStyle := p_refStyle;
      vc.col.hyperlink := nvl(p_hyperlink, false);
      
      vc.col.fmla.sharedIdx := currentCtx.sheetDefinitionMap(p_sheetId).sharedFmlaSeq;
      currentCtx.sheetDefinitionMap(p_sheetId).sharedFmlaSeq := vc.col.fmla.sharedIdx + 1;
    
    else
      -- defaults to a formula-less number column
      vc.col.type := dbms_sql.NUMBER_TYPE;
      vc.col.supertype := ST_NUMBER;
    end if;
    
    vc.pos := p_columnId;
    vc.after := p_after;
    
    t.sqlMetadata.virtualColumns.extend;
    t.sqlMetadata.virtualColumns(t.sqlMetadata.virtualColumns.last) := vc;
    
    if vc.col.hyperlink then
      currentCtx.workbook.styles.hasHlink := true;
    end if;
    
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId) := t;
  end;

  procedure addTableColumn (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_tableId         in tableHandle
  , p_name            in varchar2
  , p_value           in varchar2
  , p_refStyle        in pls_integer default null
  )
  is
  begin
    addTableColumnImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_value, null, null, p_refStyle);
  end;
  
  procedure addTableColumnBefore (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_tableId         in tableHandle
  , p_name            in varchar2
  , p_value           in varchar2
  , p_columnId        in pls_integer
  , p_refStyle        in pls_integer default null
  )
  is
  begin
    addTableColumnImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_value, p_columnId, false, p_refStyle);
  end;

  procedure addTableColumnAfter (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_tableId         in tableHandle
  , p_name            in varchar2
  , p_value           in varchar2
  , p_columnId        in pls_integer
  , p_refStyle        in pls_integer default null
  )
  is
  begin
    addTableColumnImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_value, p_columnId, true, p_refStyle);
  end;

  procedure addTableHyperlinkColImpl (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in tableHandle
  , p_name      in varchar2
  , p_location  in varchar2
  , p_linkName  in varchar2
  , p_columnId  in pls_integer
  , p_after     in boolean
  )
  is
    fmla      varchar2(32767);
  begin
    if p_location is not null then
      fmla := 'HYPERLINK(' || enquote(p_location);
      if p_name is not null then
        fmla := fmla || ',' || enquote(p_linkName);
      end if;
      fmla := fmla || ')';
    else
      error('Location parameter cannot be null');
    end if;
    addTableColumnImpl(p_ctxId, p_sheetId, p_tableId, p_name, fmla, p_columnId, p_after, ExcelFmla.REF_R1C1, true);
  end;

  procedure addTableHyperlinkColumn (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in tableHandle
  , p_name      in varchar2
  , p_location  in varchar2
  , p_linkName  in varchar2 default null
  )
  is
  begin
    addTableHyperlinkColImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_location, p_linkName, null, null);
  end;

  procedure addTableHyperlinkColumnBefore (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_name       in varchar2
  , p_columnId   in pls_integer
  , p_location   in varchar2
  , p_linkName   in varchar2 default null
  )
  is
  begin
    addTableHyperlinkColImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_location, p_linkName, p_columnId, false);
  end;

  procedure addTableHyperlinkColumnAfter (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_name       in varchar2
  , p_columnId   in pls_integer
  , p_location   in varchar2
  , p_linkName   in varchar2 default null
  )
  is
  begin
    addTableHyperlinkColImpl(p_ctxId, p_sheetId, p_tableId, p_name, p_location, p_linkName, p_columnId, true);
  end;

  procedure putDefinedName (
    p_ctxId     in ctxHandle
  , p_name      in varchar2
  , p_value     in varchar2
  , p_scope     in sheetHandle default null
  , p_comment   in varchar2 default null
  , p_cellRef   in varchar2 default null
  , p_refStyle  in pls_integer default null
  )
  is
  begin
    loadContext(p_ctxId);
    putNameImpl(
      ctxId     => p_ctxId
    , name      => p_name
    , value     => p_value
    , sheetName => case when p_scope is not null then currentCtx.sheetDefinitionMap(p_scope).sheetName end
    , cellRef   => p_cellRef
    , comment   => p_comment
    , refStyle  => p_refStyle
    );
  end;

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
  )
  is
    dvRule     ExcelTypes.CT_DataValidation;
    dvRuleIdx  pls_integer;
  begin  
  
    dvRule := makeDVRule(
                p_type
              , p_cellRange
              , p_value1
              , p_value2
              , p_operator
              , p_allowBlank
              , p_showDropDown
              , p_showErrorMessage
              , p_errorMsg
              , p_errorTitle
              , p_errorStyle
              , p_showInputMessage
              , p_promptMsg
              , p_promptTitle
              , p_refStyle1
              , p_refStyle2
              );
    
    if dvRule.sqref.ranges.count = 0 then
      error('Missing cell reference in data validation rule');
    else
      loadContext(p_ctxId);
      currentCtx.sheetDefinitionMap(p_sheetId).dvRules.extend;
      dvRuleIdx := currentCtx.sheetDefinitionMap(p_sheetId).dvRules.last;
      currentCtx.sheetDefinitionMap(p_sheetId).dvRules(dvRuleIdx) := dvRule;      
    end if;
    
  end;

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
  )
  is
    rule     ExcelTypes.CT_CfRule;
    ruleIdx  pls_integer;
  begin
    
    rule := makeCfRule(
              p_type
            , p_cellRange
            , p_style
            , p_operator
            , p_value1
            , p_value2
            , p_param
            , p_percent
            , p_cfvoList
            , p_hideValue
            , p_iconSet
            , p_reverse
            , p_stopIfTrue
            , p_refStyle1
            , p_refStyle2
            );

    loadContext(p_ctxId);
    currentCtx.sheetDefinitionMap(p_sheetId).cfRules.extend;
    ruleIdx := currentCtx.sheetDefinitionMap(p_sheetId).cfRules.last;
    currentCtx.sheetDefinitionMap(p_sheetId).cfRules(ruleIdx) := rule;

  end;

  procedure putCellImpl (
    ctxId           in ctxHandle
  , sheetId         in sheetHandle
  , rowIdx          in pls_integer
  , colIdx          in pls_integer
  , data            in data_t
  , xfId            in pls_integer
  , anchorTableId   in tableHandle default null
  , anchorPosition  in pls_integer default null
  , refStyle        in pls_integer default null
  , hyperlink       in boolean default false
  )
  is
    cell   cell_t;
    cell2  floatingCell_t;
    idx    pls_integer;
  begin
    loadContext(ctxId);
    if hyperlink then
      currentCtx.workbook.styles.hasHlink := true;
    end if;
    if anchorTableId is null then
      assertPositive(rowIdx, 'The cell row must be a positive integer.');
      assertPositive(colIdx, 'The cell column must be a positive integer.');
      cell.r := rowIdx;
      cell.cn := colIdx;
      cell.c := base26encode(cell.cn);
      cell.xfId := nvl(xfId, 0);
      cell.v := data;
      cell.hyperlink := hyperlink;
      cell.isTableCell := false;
      if data.st = ST_FORMULA then
        cell.f.expr := cell.v.varchar2_value;
        cell.f.shared := false;
        cell.f.refStyle := refStyle;
      end if;
      currentCtx.sheetDefinitionMap(sheetId).data.rows(rowIdx).id := rowIdx;
      currentCtx.sheetDefinitionMap(sheetId).data.rows(rowIdx).cells(colIdx) := cell; 
    else   
      assertTableExists(sheetId, anchorTableId);
      cell2.data := data;
      cell2.xfId := nvl(xfId, 0);
      cell2.anchorRef.tableId := anchorTableId;
      cell2.anchorRef.anchorPosition := anchorPosition;
      cell2.anchorRef.rowOffset := rowIdx;
      cell2.anchorRef.colOffset := colIdx;
      if data.st = ST_FORMULA then
        cell2.fmla.expr := cell2.data.varchar2_value;
        cell2.fmla.shared := false;
        cell2.fmla.refStyle := refStyle;
      end if;
      cell2.hyperlink := hyperlink;
      currentCtx.sheetDefinitionMap(sheetId).floatingCells.extend;
      idx := currentCtx.sheetDefinitionMap(sheetId).floatingCells.last;
      currentCtx.sheetDefinitionMap(sheetId).floatingCells(idx) := cell2; 
    end if;
  end;

  procedure putNumberCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in number
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  )
  is
    data  data_t;
  begin
    prepareNumberValue(data, p_value);
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition);
  end;

  procedure putStringCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in varchar2
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  )
  is
    data  data_t;
  begin
    prepareStringValue(data, p_value);
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition);
  end;

  procedure putDateCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in date
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  )
  is
    data  data_t;
  begin
    prepareDateValue(data, p_value);
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition);
  end;

  procedure putRichTextCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in varchar2
  , p_style           in cellStyleHandle default null 
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  )
  is
    data  data_t;
  begin
    data.xml_value := xmltype('<root>'||p_value||'</root>');
    data.st := ST_RICHTEXT;
    data.db_type := dbms_sql.VARCHAR2_TYPE;
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition);
  exception
    when xml_parse_exception then
      error('Invalid XHTML fragment');
  end;

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
  )
  is
    data  data_t;
  begin
    data.st := ST_FORMULA;
    data.varchar2_value := p_value;
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition, p_refStyle);
  end;

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
  )
  is
    fmla  varchar2(32767);
    data  data_t;
  begin
    if p_location is not null then
      fmla := 'HYPERLINK(' || enquote(p_location);
      if p_linkName is not null then
        fmla := fmla || ',' || enquote(p_linkName);
      end if;
      fmla := fmla || ')';
    else
      error('Location parameter cannot be null');
    end if;
    
    data.st := ST_FORMULA;
    data.varchar2_value := fmla;
    
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition, hyperlink => true);
    
  end;
  
  procedure putImageCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_image           in blob
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null
  )
  is
    data  data_t;
  begin
    data.st := ST_IMAGE;
    data.blob_value := p_image;
    putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, null, p_anchorTableId, p_anchorPosition);
  end;  
  
  procedure putCell (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_rowIdx          in pls_integer
  , p_colIdx          in pls_integer
  , p_value           in anydata default null
  , p_style           in cellStyleHandle default null
  , p_anchorTableId   in tableHandle default null
  , p_anchorPosition  in pls_integer default null  
  )
  is
    data  data_t;
  begin
    if p_value is not null then
      prepareData(data, p_value);
      putCellImpl(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, data, p_style, p_anchorTableId, p_anchorPosition);    
    else
      putNumberCell(p_ctxId, p_sheetId, p_rowIdx, p_colIdx, null, p_style, p_anchorTableId, p_anchorPosition);
    end if;
  end;

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
  )
  is
    INDEXTYPE_ROW  constant pls_integer := 0;
    INDEXTYPE_COL  constant pls_integer := 1;
     
    dr  CT_DrawingAnchor;
    
    function validateIndex (value in pls_integer, coordType in pls_integer) return pls_integer
    is
    begin
      if value is null then
        error('Cell coordinate cannot be NULL');
      elsif coordType = INDEXTYPE_ROW and value not between 1 and MAX_ROW_NUMBER then
        error(RANGE_INVALID_ROW, value);
      elsif coordType = INDEXTYPE_COL and value not between 1 and MAX_COLUMN_NUMBER then
        error(RANGE_INVALID_COL, value);
      end if;
      return (value - 1); -- transform into 0-based index
    end;
    
    -- validate offset value as a ST_Coordinate
    /*
    ECMA-376-1
    20.1.10.16 ST_Coordinate (Coordinate)
    This simple type represents a one dimensional position or length as either:
    - EMUs.
    - A number followed immediately by a unit identifier.
    This simple type is a union of the following types:
    - The ST_CoordinateUnqualified simple type (�20.1.10.19).
    - The ST_UniversalMeasure simple type (�22.9.2.15).
    */    
    function validateMeasure (value in varchar2, errorOnNull boolean default false) return varchar2
    is
    begin
      if errorOnNull and value is null then
        error('Measurement cannot be null');
      end if;
      -- 20.1.10.19 ST_CoordinateUnqualified
      if regexp_like(value, '^-?[0-9]+$') then
        -- Valid EMU range
        if to_number(value) not between -27273042329600 and 27273042316900 then
          error('Invalid EMU value: %s', value);
        end if;
      -- 22.9.2.15 ST_UniversalMeasure (Universal Measurement)
      -- Note: ECMA-316 does not include the pixel unit (px) but Excel actually supports it.
      elsif not regexp_like(value, '^-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi|px)$') then
        error('Invalid measurement: %s', value);
      end if;
      return value;
    end;    
    
  begin
    loadContext(p_ctxId);
  
    case p_anchorType
    when TWOCELL_ANCHOR then
      dr."from".col    := validateIndex(p_fromCol, INDEXTYPE_COL);
      dr."from".colOff := validateMeasure(nvl(p_fromColOff,'0'));
      dr."from"."row"  := validateIndex(p_fromRow, INDEXTYPE_ROW);
      dr."from".rowOff := validateMeasure(nvl(p_fromRowOff,'0'));
      dr."to".col      := validateIndex(p_toCol, INDEXTYPE_COL);
      dr."to".colOff   := validateMeasure(nvl(p_toColOff,'0'));
      dr."to"."row"    := validateIndex(p_toRow, INDEXTYPE_ROW);
      dr."to".rowOff   := validateMeasure(nvl(p_toRowOff,'0'));
      dr.editAs := nvl(p_imageProps, MOVE_RESIZE);
    
    when ONECELL_ANCHOR then
      dr."from".col    := validateIndex(p_fromCol, INDEXTYPE_COL);
      dr."from".colOff := validateMeasure(nvl(p_fromColOff,'0'));
      dr."from"."row"  := validateIndex(p_fromRow, INDEXTYPE_ROW);
      dr."from".rowOff := validateMeasure(nvl(p_fromRowOff,'0'));
      dr.ext.cx        := validateMeasure(p_extX);
      dr.ext.cy        := validateMeasure(p_extY);
      
    when ABSOLUTE_ANCHOR then
      dr.pos.x  := validateMeasure(p_posX, true);
      dr.pos.y  := validateMeasure(p_posY, true);
      dr.ext.cx := validateMeasure(p_extX);
      dr.ext.cy := validateMeasure(p_extY);
            
    else
      error('Invalid drawing anchor type: %d', p_anchorType);
    end case;
    
    dr.anchorType := p_anchorType;
    dr.content := p_image;
    
    currentCtx.sheetDefinitionMap(p_sheetId).drawing.extend;
    dr.id := currentCtx.sheetDefinitionMap(p_sheetId).drawing.last;
    currentCtx.sheetDefinitionMap(p_sheetId).drawing(dr.id) := dr;
    
  end;

  procedure setSheetProperties (
    p_ctxId                in ctxHandle
  , p_sheetId              in sheetHandle
  , p_activePaneAnchorRef  in varchar2 default null
  , p_showGridLines        in boolean default null
  , p_showRowColHeaders    in boolean default null
  , p_defaultRowHeight     in number default null
  , p_background           in blob default null
  )
  is
    cellRef  cell_ref_t;
  begin
    loadContext(p_ctxId);
    if p_activePaneAnchorRef is not null then
      cellRef := parseRangeExpr(p_activePaneAnchorRef).start_ref;
      currentCtx.sheetDefinitionMap(p_sheetId).activePaneAnchorRef := cellRef;
    end if;
    if p_defaultRowHeight is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).defaultRowHeight := p_defaultRowHeight;
    end if;
    if p_showGridLines is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).showGridLines := p_showGridLines;
    end if;
    if p_showRowColHeaders is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).showRowColHeaders := p_showRowColHeaders;
    end if;
    if p_background is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).background := p_background;
    end if;
    currentCtx.sheetDefinitionMap(p_sheetId).hasProps := ( p_activePaneAnchorRef is not null 
                                                           or p_showGridLines is not null 
                                                           or p_showRowColHeaders is not null );
  end;

  procedure setRangeStyleImpl (
    ctx        in out nocopy context_t
  , sheetId    in pls_integer
  , cellRange  in cellRange_t
  )
  is
    idx pls_integer;
  begin
    ctx.sheetDefinitionMap(sheetId).cellRanges.extend;
    idx := ctx.sheetDefinitionMap(sheetId).cellRanges.last;
    ctx.sheetDefinitionMap(sheetId).cellRanges(idx) := cellRange;
  end;
  
  procedure setRangeStyle (
    p_ctxId           in ctxHandle
  , p_sheetId         in sheetHandle
  , p_range           in varchar2
  , p_style           in cellStyleHandle
  , p_outsideBorders  in boolean default false
  )
  is
    range      range_t := parseRangeExpr(p_range);
    cellRange  cellRange_t;
  begin
    loadContext(p_ctxId);
    
    cellRange.span.anchorRef.rowOffset := range.start_ref.r;
    cellRange.span.anchorRef.colOffset := range.start_ref.cn;
    cellRange.span.rowSpan := range.end_ref.r - range.start_ref.r + 1;
    cellRange.span.colSpan := range.end_ref.cn - range.start_ref.cn + 1;
    
    cellRange.xfId := p_style;
    cellRange.outsideBorders := nvl(p_outsideBorders, false);
    
    setRangeStyleImpl(currentCtx, p_sheetId, cellRange);
  end;

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
  )
  is
    cellRange  cellRange_t;
  begin
    loadContext(p_ctxId);

    if p_anchorTableId is not null then
      assertTableExists(p_sheetId, p_anchorTableId);
    else
      assertPositive(p_rowOffset, 'The range row offset must be a positive integer.');
      assertPositive(p_colOffset, 'The range column offset must be a positive integer.');
    end if;
        
    cellRange.span.anchorRef.rowOffset := p_rowOffset;
    cellRange.span.anchorRef.colOffset := p_colOffset;
    cellRange.span.anchorRef.tableId := p_anchorTableId;
    cellRange.span.anchorRef.anchorPosition := p_anchorPosition;
    cellRange.span.rowSpan := p_rowSpan;
    cellRange.span.colSpan := p_colSpan;
    
    cellRange.xfId := p_style;
    cellRange.outsideBorders := nvl(p_outsideBorders, false);
    
    setRangeStyleImpl(currentCtx, p_sheetId, cellRange);
  end;

  procedure mergeCellsImpl (
    ctx       in out nocopy context_t
  , sheetId   in pls_integer
  , cellSpan  in cellSpan_t
  , xfId      in pls_integer
  )
  is
    idx        pls_integer;
    cellRange  cellRange_t;
  begin

    ctx.sheetDefinitionMap(sheetId).mergedCells.extend;
    idx := ctx.sheetDefinitionMap(sheetId).mergedCells.last;
    ctx.sheetDefinitionMap(sheetId).mergedCells(idx) := cellSpan;

    if xfId != 0 then
      cellRange.span := cellSpan;
      cellRange.xfId := xfId;
      cellRange.outsideBorders := true;
      setRangeStyleImpl(ctx, sheetId, cellRange);
    end if;
    
  end;

  procedure mergeCells (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_range    in varchar2
  , p_style    in cellStyleHandle default null
  )
  is
    range     range_t := parseRangeExpr(p_range);
    cellSpan  cellSpan_t;
  begin
    loadContext(p_ctxId);
    
    cellSpan.anchorRef.rowOffset := range.start_ref.r;
    cellSpan.anchorRef.colOffset := range.start_ref.cn;
    cellSpan.rowSpan := range.end_ref.r - range.start_ref.r + 1;
    cellSpan.colSpan := range.end_ref.cn - range.start_ref.cn + 1;
    
    mergeCellsImpl(currentCtx, p_sheetId, cellSpan, p_style);
  end;

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
  )
  is
    cellSpan  cellSpan_t;
  begin
    loadContext(p_ctxId);
    
    if p_anchorTableId is not null then
      assertTableExists(p_sheetId, p_anchorTableId);
    else
      assertPositive(p_rowOffset, 'The range row offset must be a positive integer.');
      assertPositive(p_colOffset, 'The range column offset must be a positive integer.');
    end if;
    
    cellSpan.anchorRef.rowOffset := p_rowOffset;
    cellSpan.anchorRef.colOffset := p_colOffset;
    cellSpan.anchorRef.tableId := p_anchorTableId;
    cellSpan.anchorRef.anchorPosition := p_anchorPosition;
    cellSpan.rowSpan := p_rowSpan;
    cellSpan.colSpan := p_colSpan;
    
    mergeCellsImpl(currentCtx, p_sheetId, cellSpan, p_style);
  end;

  procedure setTableHeader (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_tableId     in tableHandle
  , p_style       in cellStyleHandle default null
  , p_autoFilter  in boolean default false
  )
  is
    tableHeader  table_header_t;
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    tableHeader.show := true;
    --tableHeader.xfId := p_style;
    tableHeader.autoFilter := p_autoFilter;
    
    if p_style is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).rowMap(0).xfId := p_style;
    end if;
    
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).header := tableHeader;
  end;

  -- DEPRECATED
  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetName   in varchar2
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  )
  is
  begin
    loadContext(p_ctxId);
    setHeader(p_ctxId, currentCtx.sheetIndexMap(upper(p_sheetName)), p_style, p_frozen, p_autofilter);
  end;

  procedure setHeader (
    p_ctxId       in ctxHandle
  , p_sheetId     in sheetHandle
  , p_style       in cellStyleHandle default null
  , p_frozen      in boolean default false
  , p_autoFilter  in boolean default false
  )
  is
    tableId         pls_integer;
    tableAnchorRef  anchorRef_t;
  begin
    loadContext(p_ctxId);
    tableId := currentCtx.sheetDefinitionMap(p_sheetId).tableList.first;
    setTableHeader(p_ctxId, p_sheetId, tableId, p_style, p_autofilter);
    if p_frozen then
      -- make header of first table frozen
      tableAnchorRef := currentCtx.sheetDefinitionMap(p_sheetId).tableList(tableId).anchorRef;    
      currentCtx.sheetDefinitionMap(p_sheetId).activePaneAnchorRef := makeCellRef('A', tableAnchorRef.rowOffset + 1);
      currentCtx.sheetDefinitionMap(p_sheetId).hasProps := true;
    end if;
  end;

  procedure setTableRowProperties (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_tableId  in pls_integer
  , p_rowId    in pls_integer default null
  , p_style    in cellStyleHandle default null
  , p_height   in number default null
  )
  is
    props  rowProperties_t;
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    props.xfId := p_style;
    props.height := p_height;
    if p_rowId is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).rowMap(p_rowId) := props;
    else
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).rowProps := props;
    end if;
  end;

  procedure setTableColumnProperties (
    p_ctxId        in ctxHandle
  , p_sheetId      in sheetHandle
  , p_tableId      in pls_integer
  , p_columnId     in pls_integer
  , p_columnName   in varchar2 default null
  , p_style        in cellStyleHandle default null
  , p_headerStyle  in cellStyleHandle default null
  )
  is
    tableColumn  table_column_t;
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    if currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap.exists(p_columnId) then
      tableColumn := currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId);
    end if;
    tableColumn.name := p_columnName;
    tableColumn.xfId := p_style;
    tableColumn.headerXfId := p_headerStyle;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId) := tableColumn;
  end;

  procedure setTableColumnFormat (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_tableId   in pls_integer
  , p_columnId  in pls_integer
  , p_format    in varchar2
  )
  is
    xfId         pls_integer;
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    -- get existing xfId for this column
    if currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap.exists(p_columnId) then
      xfId := currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).xfId;
    end if;
    xfId := mergeCellFormat(currentCtx, xfId, p_format, force => true);    
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).xfId := xfId;
  end;

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
  )
  is
    dvRule  ExcelTypes.CT_DataValidation;
  begin
    dvRule := makeDVRule(
                p_type
              , null
              , p_value1
              , p_value2
              , p_operator
              , p_allowBlank
              , p_showDropDown
              , p_showErrorMessage
              , p_errorMsg
              , p_errorTitle
              , p_errorStyle
              , p_showInputMessage
              , p_promptMsg
              , p_promptTitle
              , p_refStyle1
              , p_refStyle2
              );

    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).dvRule := dvRule; 
  end;

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
  )
  is
    rule     ExcelTypes.CT_CfRule;
    ruleIdx  pls_integer;
  begin
    
    rule := makeCfRule(
              p_type
            , null
            , p_style
            , p_operator
            , p_value1
            , p_value2
            , p_param
            , p_percent
            , p_cfvoList
            , p_hideValue
            , p_iconSet
            , p_reverse
            , p_stopIfTrue
            , p_refStyle1
            , p_refStyle2
            );

    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    
    if p_columnId is not null then
      
      -- column-level formatting
      if not currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap.exists(p_columnId)
        or currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).cfRules is null 
      then
        currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).cfRules := ExcelTypes.CT_CfRules();
      end if;
      
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).cfRules.extend;
      ruleIdx := currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).cfRules.last;
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).columnMap(p_columnId).cfRules(ruleIdx) := rule;

    else
      
      -- table(row)-level formatting
      if currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).cfRules is null then
        currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).cfRules := ExcelTypes.CT_CfRules();
      end if;
      
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).cfRules.extend;
      ruleIdx := currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).cfRules.last;
      currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).cfRules(ruleIdx) := rule;
    
    end if;

  end;

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
  )
  is
    nameKey  varchar2(2048);
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetId, p_tableId);
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).formatAsTable := true;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).tableStyle := p_style;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).showFirstColumn := p_showFirstColumn;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).showLastColumn := p_showLastColumn;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).showRowStripes := p_showRowStripes;
    currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).showColumnStripes := p_showColumnStripes;
    
    if p_tableName is not null then
      nameKey := upper(p_tableName);
      -- table and defined names must be unique
      if currentCtx.nameMap.exists(nameKey) then
        error('Name already exists: %s', p_tableName);
      else
        -- add a map entry for this name
        currentCtx.nameMap(nameKey) := null;
        currentCtx.sheetDefinitionMap(p_sheetId).tableList(p_tableId).tableName := p_tableName;
      end if;
    end if;
    
  end;

  -- DEPRECATED
  procedure setTableFormat (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_style      in varchar2 default null
  )
  is
  begin
    loadContext(p_ctxId);
    setTableFormat(p_ctxId, currentCtx.sheetIndexMap(upper(p_sheetName)), p_style);
  end;

  procedure setTableFormat (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_style      in varchar2 default null
  )
  is
    tableId  pls_integer;
  begin
    loadContext(p_ctxId);
    tableId := currentCtx.sheetDefinitionMap(p_sheetId).tableList.first;
    setTableProperties(p_ctxId, p_sheetId, tableId, p_style);
  end;

  procedure setDateFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.defaultFmts.dateFmt := p_format;
  end;

  procedure setDateFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.sheetDefinitionMap(p_sheetId).defaultFmts.dateFmt := p_format;
  end;

  procedure setNumFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.defaultFmts.numFmt := p_format;
  end;

  procedure setNumFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.sheetDefinitionMap(p_sheetId).defaultFmts.numFmt := p_format;
  end;

  procedure setTimestampFormat (
    p_ctxId   in ctxHandle
  , p_format  in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.defaultFmts.timestampFmt := p_format;
  end;

  procedure setTimestampFormat (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_format   in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.sheetDefinitionMap(p_sheetId).defaultFmts.timestampFmt := p_format;
  end;

  procedure setRowProperties (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_rowId    in pls_integer
  , p_style    in cellStyleHandle default null
  , p_height   in number default null
  )
  is
    r  row_t;
  begin
    loadContext(p_ctxId);
    begin
      r := currentCtx.sheetDefinitionMap(p_sheetId).data.rows(p_rowId);
    exception
      when no_data_found then
        r.id := p_rowId;
    end;
    r.props.xfId := p_style;
    r.props.height := p_height;    
    currentCtx.sheetDefinitionMap(p_sheetId).data.rows(r.id) := r;
  end;

  procedure setColumnProperties (
    p_ctxId     in ctxHandle
  , p_sheetId   in sheetHandle
  , p_columnId  in pls_integer
  , p_style     in cellStyleHandle default null
  , p_header    in varchar2 default null
  , p_width     in number default null
  )
  is
    props    colProperties_t;
    tableId  pls_integer;
  begin
    loadContext(p_ctxId);
    
    props.xfId := p_style;
    props.width := p_width;
    
    if p_header is not null then
      tableId := currentCtx.sheetDefinitionMap(p_sheetId).tableList.first;
      if tableId is not null then
        setTableColumnProperties(p_ctxId, p_sheetId, tableId, p_columnId, p_header);
      end if;
    end if;
    
    if p_width is not null or p_style is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).hasCustomColProps := true;
    end if;
    
    currentCtx.sheetDefinitionMap(p_sheetId).columnMap(p_columnId) := props;
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
    xfId    pls_integer;
  begin
    loadContext(p_ctxId);
    -- get existing xfId for this column
    if currentCtx.sheetDefinitionMap(p_sheetId).columnMap.exists(p_columnId) then
      xfId := currentCtx.sheetDefinitionMap(p_sheetId).columnMap(p_columnId).xfId;
    end if;
    xfId := mergeCellFormat(currentCtx, xfId, p_format, force => true);    
    setColumnProperties(p_ctxId, p_sheetId, p_columnId, xfId, p_header, p_width);
  end;

  procedure setBindVariableImpl (
    p_ctxId       in ctxHandle
  , p_sheetIndex  in pls_integer
  , p_bindName    in varchar2
  , p_bindValue   in anydata
  , p_tableId     in pls_integer
  )
  is
    bindVarList  bind_variable_list_t;
    varIdx       pls_integer;
  begin
    loadContext(p_ctxId);
    assertTableExists(p_sheetIndex, p_tableId);
    bindVarList := currentCtx.sheetDefinitionMap(p_sheetIndex).tableList(p_tableId).sqlMetadata.bindVariables;
    bindVarList.extend;
    varIdx := bindVarList.last;
    bindVarList(varIdx).name := p_bindName;
    bindVarList(varIdx).value := p_bindValue;     
    currentCtx.sheetDefinitionMap(p_sheetIndex).tableList(p_tableId).sqlMetadata.bindVariables := bindVarList;
  end;

  procedure setDefaultStyle (
    p_ctxId  in ctxHandle
  , p_style  in cellStyleHandle
  )
  is
  begin
    loadContext(p_ctxId);
    if p_style is not null then
      currentCtx.defaultXfId := p_style;
    end if;
  end;

  procedure setDefaultStyle (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle
  , p_style    in cellStyleHandle
  )
  is
  begin
    loadContext(p_ctxId);
    if p_sheetId is not null and p_style is not null then
      currentCtx.sheetDefinitionMap(p_sheetId).defaultXfId := p_style;
    end if;
  end;

  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in number
  )
  is
  begin
    loadContext(p_ctxId);
    setBindVariableImpl(p_ctxId, currentCtx.sheetIndexMap(upper(p_sheetName)), p_bindName, anydata.ConvertNumber(p_bindValue), 1);
  end;
  
  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  )
  is
  begin
    loadContext(p_ctxId);
    setBindVariableImpl(p_ctxId, currentCtx.sheetIndexMap(upper(p_sheetName)), p_bindName, anydata.ConvertVarchar2(p_bindValue), 1);
  end;

  -- DEPRECATED
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetName  in varchar2
  , p_bindName   in varchar2
  , p_bindValue  in date
  )
  is
  begin
    loadContext(p_ctxId);
    setBindVariableImpl(p_ctxId, currentCtx.sheetIndexMap(upper(p_sheetName)), p_bindName, anydata.ConvertDate(p_bindValue), 1);
  end;
    
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in number
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertNumber(p_bindValue), p_tableId);
  end;
  
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertVarchar2(p_bindValue), p_tableId);
  end;

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_tableId    in tableHandle
  , p_bindName   in varchar2
  , p_bindValue  in date
  )
  is
  begin
    setBindVariableImpl(p_ctxId, p_sheetId, p_bindName, anydata.ConvertDate(p_bindValue), p_tableId);
  end;

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in number
  )
  is
  begin
    setBindVariable(p_ctxId, p_sheetId, 1, p_bindName, p_bindValue);
  end;
  
  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in varchar2
  )
  is
  begin
    setBindVariable(p_ctxId, p_sheetId, 1, p_bindName, p_bindValue);
  end;

  procedure setBindVariable (
    p_ctxId      in ctxHandle
  , p_sheetId    in sheetHandle
  , p_bindName   in varchar2
  , p_bindValue  in date
  )
  is
  begin
    setBindVariable(p_ctxId, p_sheetId, 1, p_bindName, p_bindValue);
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
    loadContext(p_ctxId);
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
    
    currentCtx.encryptionInfo := encInfo;
      
  end;
$end

  procedure setCellReferenceStyle (
    p_ctxId     in ctxHandle
  , p_refStyle  in pls_integer
  )
  is
  begin
    loadContext(p_ctxId);
    if p_refStyle is null or p_refStyle not in (ExcelFmla.REF_A1, ExcelFmla.REF_R1C1) then
      error('Invalid cell reference style');
    end if;
    currentCtx.workbook.refStyle := p_refStyle;
  end;

  procedure setCoreProperties (
    p_ctxId        in ctxHandle
  , p_creator      in varchar2 default null
  , p_description  in varchar2 default null
  , p_subject      in varchar2 default null
  , p_title        in varchar2 default null
  )
  is
  begin
    loadContext(p_ctxId);
    currentCtx.coreProperties.creator := p_creator;
    currentCtx.coreProperties.description := p_description;
    currentCtx.coreProperties.subject := p_subject;
    currentCtx.coreProperties.title := p_title;
  end;

  procedure createCoreProperties (
    ctx  in out nocopy context_t 
  )
  is
    stream  stream_t := new_stream();
  begin
    stream_write(stream, '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
    stream_write(stream, '<dcterms:created xsi:type="dcterms:W3CDTF">'||to_char(systimestamp at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS"Z"')||'</dcterms:created>');
    stream_write(stream, '<dc:creator>'||dbms_xmlgen.convert(nvl(ctx.coreProperties.creator, getProductName()))||'</dc:creator>');
    stream_write(stream, '<dc:description>'||dbms_xmlgen.convert(ctx.coreProperties.description)||'</dc:description>');
    stream_write(stream, '<dc:subject>'||dbms_xmlgen.convert(ctx.coreProperties.subject)||'</dc:subject>');
    stream_write(stream, '<dc:title>'||dbms_xmlgen.convert(ctx.coreProperties.title)||'</dc:title>');
    stream_write(stream, '</cp:coreProperties>');
    stream_flush(stream);
    addPart(ctx, 'docProps/core.xml', MT_CORE, stream.content);
    addRelationship(ctx, null, RS_CORE, 'docProps/core.xml');
  end;

  procedure createMetadata (
    ctx  in out nocopy context_t
  )
  is
    rdRichValueStructure  part_t;
    rdRichValueStructureStream  stream_t;
    
    richValueRel        part_t;
    richValueRelStream  stream_t;
    
    rdRichValue         part_t;
    rdRichValueStream   stream_t;
    
    metadata            part_t;
    metadataStream      stream_t;
    metadataBinStream   xutl_xlsb.stream_t;
    
    rId           varchar2(256);
    structId      pls_integer;
    imageCount    pls_integer := ctx.imageMetadata.partMap.count;
    
  begin
    
    if imageCount != 0 then
  
      structId := putLocalImageRvStruct(ctx.rvStructures);
      
      -- rdRichValueStructure.xml
      rdRichValueStructure := new_part('xl/richData/rdrichvaluestructure.xml', MT_RDRICHVALUESTRUCTURE);
      rdRichValueStructureStream := new_stream();
      stream_write(rdRichValueStructureStream, '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="'||to_char(ctx.rvStructures.structs.count)||'">');
      for i in 1 .. ctx.rvStructures.structs.count loop
        stream_write(rdRichValueStructureStream, '<s t="'||ctx.rvStructures.structs(i).t||'">');
        for j in 1 .. ctx.rvStructures.structs(i).keys.count loop
          stream_write(rdRichValueStructureStream, '<k n="'||ctx.rvStructures.structs(i).keys(j).n||'" t="'||ctx.rvStructures.structs(i).keys(j).t||'"/>');
        end loop;
        stream_write(rdRichValueStructureStream, '</s>');
      end loop;
      stream_write(rdRichValueStructureStream, '</rvStructures>');
      stream_flush(rdRichValueStructureStream);
      rdRichValueStructure.content := rdRichValueStructureStream.content;
      addPart(ctx, rdRichValueStructure);
      addRelationship(ctx, ctx.workbook.partName, RS_RDRICHVALUESTRUCTURE, rdRichValueStructure.name);
      
      
      -- richValueRel.xml
      -- this part holds relationship ids to image parts
      richValueRel := new_part('xl/richData/richValueRel.xml', MT_RICHVALUEREL);
      richValueRelStream := new_stream();
      stream_write(richValueRelStream, '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
      
      for i in 1 .. imageCount loop 
        rId := addRelationship(richValueRel, RS_IMAGE, ctx.imageMetadata.partMap(i));
        stream_write(richValueRelStream, '<rel r:id="'||rId||'"/>');   
      end loop;
      
      stream_write(richValueRelStream, '</richValueRels>');
      stream_flush(richValueRelStream);    
      richValueRel.content := richValueRelStream.content;
      addPart(ctx, richValueRel);
      addRelationship(ctx, ctx.workbook.partName, RS_RICHVALUEREL, richValueRel.name);
      
      
      -- rdRichValue.xml
      -- the first value of each rv record is the 0-based index of the image in the richValueRel part 
      rdRichValue := new_part('xl/richData/rdrichvalue.xml', MT_RDRICHVALUE);
      rdRichValueStream := new_stream();
      
      stream_write(rdRichValueStream, '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="'||to_char(imageCount)||'">');
      for i in 1 .. imageCount loop
        stream_write(rdRichValueStream, '<rv s="'||to_char(structId)||'"><v>'||to_char(i-1)||'</v><v>5</v></rv>');
      end loop;
      
      stream_write(rdRichValueStream, '</rvData>');
      stream_flush(rdRichValueStream);
      rdRichValue.content := rdRichValueStream.content;
      addPart(ctx, rdRichValue);
      addRelationship(ctx, ctx.workbook.partName, RS_RDRICHVALUE, rdRichValue.name);
      
      if ctx.fileType = FILE_XLSX then
      
        -- metadata.xml
        metadata := new_part('xl/metadata.xml', MT_METADATA);
        metadataStream := new_stream();
        stream_write(metadataStream, '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">');
        stream_write(metadataStream, '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>');
        
        stream_write(metadataStream, '<futureMetadata name="XLRICHVALUE" count="'||to_char(imageCount)||'">');
        -- 0-based index of a richValue (rv) block in rdRichValue.xml part
        for i in 1 .. imageCount loop
          stream_write(metadataStream, '<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="'||to_char(i-1)||'"/></ext></extLst></bk>');
        end loop;
        stream_write(metadataStream, '</futureMetadata>');  
        
        stream_write(metadataStream, '<valueMetadata count="'||to_char(imageCount)||'">');
        -- 0-based index of a futureMetadata block
        for i in 1 .. imageCount loop        
          stream_write(metadataStream, '<bk><rc t="1" v="'||to_char(i-1)||'"/></bk>');
        end loop; 
        stream_write(metadataStream, '</valueMetadata>');
        
        stream_write(metadataStream, '</metadata>');
        stream_flush(metadataStream); --TODO? include flush flag in stream_write    
        metadata.content := metadataStream.content; -- TODO? include stream_t in part_t
      
      else
        
        -- metadata.bin
        metadata := new_part('xl/metadata.bin', MT_METADATA_BIN);
        metadata.isBinary := true;
        metadataBinStream := xutl_xlsb.new_stream();
        xutl_xlsb.put_Metadata(metadataBinStream, imageCount);
        xutl_xlsb.flush_stream(metadataBinStream);
        metadata.contentBin := metadataBinStream.content;
        
      end if;
      
      addPart(ctx, metadata);
      addRelationship(ctx, ctx.workbook.partName, RS_METADATA, metadata.name);
    
    end if;
    
  end;

  function getFileContent (
    p_ctxId  in ctxHandle
  )
  return blob
  is
    sh        sheetHandle;
    shHandles  intList_t := intList_t();
    sheet      ExcelTypes.CT_SheetBase;
    sheets     ExcelTypes.CT_Sheets := ExcelTypes.CT_Sheets();
    --sd         sheet_definition_t;
    output     blob;
  begin
    loadContext(p_ctxId);
    -- shared styles
    addDefaultStyles(currentCtx.workbook.styles);
  
    -- the following loop:
    -- builds a collection of sheet handles
    -- builds a collection of (sheetName, sheetIdx) tuples to be passed to the formula context
    sh := currentCtx.sheetDefinitionMap.first;
    while sh is not null loop
      -- list of sheet handles
      shHandles.extend;
      sheet.idx := shHandles.last;
      shHandles(sheet.idx) := sh;
      
      if not currentCtx.sheetDefinitionMap(sh).pageable then
        sheet.name := currentCtx.sheetDefinitionMap(sh).sheetName;
        sheets.extend;
        sheets(sheets.last) := sheet;
        
        -- update sheet definition and index map with new densified index
        --currentCtx.sheetDefinitionMap(sh).sheetIndex := sheet.idx;
        --currentCtx.sheetIndexMap(upper(sheet.name)) := sheet.idx;
        
      end if;
      
      sh := currentCtx.sheetDefinitionMap.next(sh);
    end loop;
    
    -- formula context
    ExcelFmla.setContext(sheets, currentCtx.names);
    
    -- worksheets
    for i in 1 .. shHandles.count loop
      createWorksheet(currentCtx, shHandles(i));
    end loop;
    
    -- workbook
    case currentCtx.fileType
    when FILE_XLSX then
      createWorkbook(currentCtx);
    when FILE_XLSB then
      createWorkbookBin(currentCtx);
    end case;
    
    -- core properties
    createCoreProperties(currentCtx);
    
    -- metadata
    currentCtx.rvStructures.structs := ExcelTypes.CT_RichValueStructureList();
    createMetadata(currentCtx);
    
    createContentTypes(currentCtx);
    createRels(currentCtx);
    
    debug('start create package');  
    createPackage(currentCtx.pck);  
    debug('end create package');
    
$if NOT $$no_crypto OR $$no_crypto IS NULL $then
    if currentCtx.encryptionInfo.version is not null then
      output := xutl_offcrypto.encrypt_package(
                  p_package  => currentCtx.pck.content
                , p_password => currentCtx.encryptionInfo.password
                , p_version  => currentCtx.encryptionInfo.version
                , p_cipher   => currentCtx.encryptionInfo.cipherName
                , p_hash     => currentCtx.encryptionInfo.hashFuncName
                );
      dbms_lob.freetemporary(currentCtx.pck.content);
    else    
$end
      output := currentCtx.pck.content;
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

  function getRowCount (
    p_ctxId    in ctxHandle
  , p_sheetId  in sheetHandle 
  , p_tableId  in tableHandle default null
  ) 
  return pls_integer
  is
    tableId  pls_integer := nvl(p_tableId, currentCtx.sheetDefinitionMap(p_sheetId).tableList.first);
  begin
    loadContext(p_ctxId);
    return currentCtx.sheetDefinitionMap(p_sheetId).tableList(tableId).sqlMetadata.r_num;
  end;

begin
  
  init;

end ExcelGen;
/
