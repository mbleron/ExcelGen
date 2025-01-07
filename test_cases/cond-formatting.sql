declare

  ctx     ExcelGen.ctxHandle := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  sheet1  ExcelGen.sheetHandle;
  sheet2  ExcelGen.sheetHandle;
  sheet3  ExcelGen.sheetHandle;
  sheet4  ExcelGen.sheetHandle;
  table1  ExcelGen.tableHandle;
  
  whites1 varchar2(8 char) := unistr('\2659\2659\2659\2659\2659\2659\2659\2659');
  whites2 varchar2(8 char) := unistr('\2656\2658\2657\2655\2654\2657\2658\2656');
  blacks1 varchar2(8 char) := unistr('\265F\265F\265F\265F\265F\265F\265F\265F');
  blacks2 varchar2(8 char) := unistr('\265C\265E\265D\265B\265A\265D\265E\265C');

  color1  varchar2(128);
  color2  varchar2(128);
  
  col1    pls_integer;
  col2    pls_integer;
  
  function hsl2rgb (H in number, SL in number, L in number)
  return varchar2
  is
    a  number := SL * least(L, 1 - L);
    function f (n in number) return number is
      k  number := mod(n + H/30, 12);
    begin
      return L - a * greatest(-1, least(k - 3, 9 - k, 1));
    end;
  begin
    return ExcelGen.makeRgbColor(round(f(0) * 255), round(f(8) * 255), round(f(4) * 255));
  end;  
  
begin
  
  -- S H E E T 1
  -- Miscellaneous conditional formatting rules
  sheet1 := ExcelGen.addSheet(ctx, 'Misc.');
  
  for i in 1 .. 9 loop
    for j in 1 .. 5 loop
      if i = 8 and j in (1,2) then
        ExcelGen.putNumberCell(ctx, sheet1, i, j, 1);
      else
        ExcelGen.putNumberCell(ctx, sheet1, i, j, j);
      end if;
    end loop;
  end loop;
  
  ExcelGen.putNumberCell(ctx, sheet1, 11, 1, 2); -- A11
  ExcelGen.putNumberCell(ctx, sheet1, 11, 2, 4); -- B11
  ExcelGen.putStringCell(ctx, sheet1, 11, 3, 'X'); -- C11

  -- 3-color scale
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_COLORSCALE
  , p_cellRange => ExcelTypes.ST_Sqref('A1:E1')
  , p_cfvoList  => ExcelTypes.CT_CfvoList(
                     ExcelTypes.makeCfvo(ExcelTypes.CFVO_MIN, p_color => '#F8696B') -- Minimum
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_PERCENTILE, p_value => 50, p_color => '#FFEB84') -- Midpoint
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_MAX, p_color => '#63BE7B') -- Maximum
                   )
  );

  -- 2-color scale
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_COLORSCALE
  , p_cellRange => ExcelTypes.ST_Sqref('A2:E2')
  , p_cfvoList  => ExcelTypes.CT_CfvoList(
                     ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, p_value => 0, p_color => 'white') -- Minimum
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, p_value => 10, p_color => 'green') -- Maximum
                   )
  );

  -- top 1
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_TOP
  , p_cellRange => ExcelTypes.ST_Sqref('A3:E3')
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'border:medium solid green')
  , p_param     => 1
  );

  -- Icon set
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_ICONSET
  , p_cellRange => ExcelTypes.ST_Sqref('A4:E4')
  , p_iconSet   => ExcelTypes.CF_ICONSET_5QUARTERS
  , p_cfvoList  => ExcelTypes.CT_CfvoList(
                     ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, 1) -- Threshold #1
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, 2) -- Threshold #2
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, 3) -- Threshold #3
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, 4) -- Threshold #4
                   )
  );
  
  -- Cell value
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_CELLIS
  , p_cellRange => ExcelTypes.ST_Sqref('A5:E5')
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'border:thin solid black;background-color:yellow')
  , p_operator  => ExcelTypes.CF_OPER_BN
  , p_value1    => '$A$11'
  , p_value2    => '$B$11'
  );

  -- Data bar
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_DATABAR
  , p_cellRange => ExcelTypes.ST_Sqref('A6:E6')
  , p_cfvoList  => ExcelTypes.CT_CfvoList(
                     ExcelTypes.makeCfvo(ExcelTypes.CFVO_MIN) -- Minimum
                   , ExcelTypes.makeCfvo(ExcelTypes.CFVO_MAX) -- Maximum
                   , ExcelTypes.makeCfvo(p_color => 'hotpink') -- Bar color
                   )
  , p_hideValue => true 
  );

  -- Equal or Above Average
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_EQUALABOVEAVERAGE
  , p_cellRange => ExcelTypes.ST_Sqref('A7:E7')
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'background-color:orange')
  );
  
  -- Duplicate values
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_DUPLICATES
  , p_cellRange => ExcelTypes.ST_Sqref('A8:E8')
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'border:medium solid red;background-color:#FFA7A7')
  );
  
  -- Formula-based
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => ExcelTypes.CF_TYPE_EXPR
  , p_cellRange => ExcelTypes.ST_Sqref('A9:E9')
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'font-style:italic;color:#BFBFBF;background-color:#F2F2F2')
  , p_value1    => 'R11C3="X"'
  , p_refstyle1 => ExcelFmla.REF_R1C1
  );  
    
  -- S H E E T 2
  -- Chessboard
  sheet2 := ExcelGen.addSheet(ctx, 'Chessboard');
  
  for i in 1 .. 8 loop
    ExcelGen.setColumnProperties(ctx, sheet2, i+1, p_width => ExcelGen.colPxToCharWidth(40));
    ExcelGen.setRowProperties(ctx, sheet2, i+1, p_height => ExcelGen.rowPxToPt(40));
  end loop;
  
  ExcelGen.setColumnProperties(ctx, sheet2, 1, p_width => ExcelGen.colPxToCharWidth(20));
  ExcelGen.setSheetProperties(ctx, sheet2, p_showGridLines => false);
  
  for i in 1 .. 8 loop
    ExcelGen.putStringCell(ctx, sheet2, 3, i+1, substr(blacks1,i,1));
    ExcelGen.putStringCell(ctx, sheet2, 2, i+1, substr(blacks2,i,1));
    ExcelGen.putStringCell(ctx, sheet2, 8, i+1, substr(whites1,i,1));
    ExcelGen.putStringCell(ctx, sheet2, 9, i+1, substr(whites2,i,1));
  end loop;
  
  ExcelGen.setRangeStyle(ctx, sheet2, 'B2:I9', p_style => ExcelGen.makeCellStyleCss(ctx, 'text-align:center;vertical-align:middle;font-size:20pt;border:medium solid black'), p_outsideBorders => true);
  
  -- check pattern achieved via the following conditional formatting rule
  ExcelGen.addCondFormattingRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet2
  , p_cellRange => ExcelTypes.ST_Sqref('B2:I9')
  , p_type      => ExcelTypes.CF_TYPE_EXPR
  , p_style     => ExcelGen.makeCondFmtStyleCss(ctx, 'background-color:tan')
  , p_value1    => 'MOD(ROW(),2)<>MOD(COLUMN(),2)'
  );
  
  -- S H E E T 3
  -- HSL color scale
  sheet3 := ExcelGen.addSheet(ctx, 'Colorscale');
  ExcelGen.setSheetProperties(ctx, sheet3, p_showGridLines => false, p_showRowColHeaders => false);

  for i in 1 .. 101 loop
    for j in 1 .. 360 loop
      ExcelGen.putNumberCell(ctx, sheet3, i, j, j);
    end loop;
  end loop;
  
  for lt in 0 .. 100 loop
    color1 := null;
    for v in 0 .. 6 loop
      color2 := hsl2rgb(v*60, 1, lt/100);
      col2 := case when v = 0 then 1 else v*60 end;

      if color1 is not null then
        ExcelGen.addCondFormattingRule(
          ctx
        , sheet3
        , ExcelTypes.CF_TYPE_COLORSCALE
        , p_cellRange => ExcelTypes.ST_Sqref(ExcelGen.makeCellRange(lt+1, col1, lt+1, col2))
        , p_cfvoList => ExcelTypes.CT_CfvoList(
                          ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, col1, p_color => color1)
                        , ExcelTypes.makeCfvo(ExcelTypes.CFVO_NUM, col2, p_color => color2)
                        )
        );
      end if;
      col1 := col2;
      color1 := color2;
    end loop;
  end loop;
  
  for colIdx in 1 .. 360 loop
    ExcelGen.setColumnProperties(ctx, sheet3, colIdx, p_width => 3);
  end loop;
  
  -- S H E E T 4
  -- Table-level conditional formatting
  sheet4 := ExcelGen.addSheet(ctx, 'Employees');
  table1 := ExcelGen.addTable(ctx, sheet4, 'SELECT * FROM HR.EMPLOYEES');
  
  ExcelGen.setDateFormat(ctx, sheet4, 'dd/mm/yyyy');
  ExcelGen.setTableHeader(ctx, sheet4, table1, p_autoFilter => true);
  ExcelGen.setSheetProperties(ctx, sheet4, p_activePaneAnchorRef => 'A2');
  
  -- highlight salaries above average
  ExcelGen.addTableCondFmtRule(
    p_ctxId    => ctx
  , p_sheetId  => sheet4
  , p_tableId  => table1
  , p_columnId => 8 -- SALARY
  , p_type     => ExcelTypes.CF_TYPE_ABOVEAVERAGE
  , p_style    => ExcelGen.makeCondFmtStyleCss(ctx, 'background-color:lightskyblue;font-weight:bold')
  );

  -- highlight rows where employee has no manager (column #10)
  ExcelGen.addTableCondFmtRule(
    p_ctxId    => ctx
  , p_sheetId  => sheet4
  , p_tableId  => table1
  , p_columnId => null
  , p_type     => ExcelTypes.CF_TYPE_EXPR
  , p_style    => ExcelGen.makeCondFmtStyleCss(ctx, 'color:red;font-style:italic')
  , p_value1   => 'ISBLANK(RC10)'  -- MANAGER_ID
  , p_refstyle1 => ExcelFmla.REF_R1C1
  );
  
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'cond-formatting.xlsx');
  ExcelGen.closeContext(ctx);
  
end;
/
