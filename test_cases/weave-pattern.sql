declare

  ctx     ExcelGen.ctxHandle;
  sheet1  ExcelGen.sheetHandle;
  altRow  boolean := false;
  altCol  boolean := false;
  
  patternStyle1  ExcelGen.cellStyleHandle;
  patternStyle2  ExcelGen.cellStyleHandle;
  patternStyle3  ExcelGen.cellStyleHandle;
  
  procedure putBlockPattern1 (rowIdx in pls_integer, colIdx pls_integer) is
  begin
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx-1, 2*colIdx-1, p_style => patternStyle1);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx-1, 2*colIdx, p_style => patternStyle2);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx, 2*colIdx-1, p_style => patternStyle3);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx, 2*colIdx,   p_style => patternStyle3);    
  end;
  
  procedure putBlockPattern2 (rowIdx in pls_integer, colIdx pls_integer) is
  begin
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx-1, 2*colIdx-1, p_style => patternStyle1);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx-1, 2*colIdx, p_style => patternStyle2);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx, 2*colIdx-1, p_style => patternStyle3);
    ExcelGen.putCell(ctx, sheet1, 2*rowIdx, 2*colIdx,   p_style => patternStyle2);            
  end;

begin
  
  ctx := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  sheet1 := ExcelGen.addSheet(ctx, 'sheet1');
  
  patternStyle1 := ExcelGen.makeCellStyleCss(ctx, 'background:black');
  patternStyle2 := ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to right,#333333,#8ea9db 10% 80%,#333333);border-left:thin solid #444444;border-right:thin solid #444444');
  patternStyle3 := ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to bottom,#333333,#ffc000 10% 80%,#333333);border-top:thin solid #444444;border-bottom:thin solid #444444');
    
  for rowIdx in 1 .. 20 loop
    
    ExcelGen.setRowProperties(ctx, sheet1, 2*rowIdx-1, p_height => ExcelGen.rowPxToPt(8));
    ExcelGen.setRowProperties(ctx, sheet1, 2*rowIdx, p_height => ExcelGen.rowPxToPt(40));
    
    altCol := altRow;
  
    for colIdx in 1 .. 40 loop
      
      if rowIdx = 1 then
        ExcelGen.setColumnProperties(ctx, sheet1, 2*colIdx-1, p_width => ExcelGen.colPxToCharWidth(8));
        ExcelGen.setColumnProperties(ctx, sheet1, 2*colIdx, p_width => ExcelGen.colPxToCharWidth(40));
      end if;
      
      if not altCol then
        putBlockPattern1(rowIdx, colIdx);
      else
        putBlockPattern2(rowIdx, colIdx);
      end if;
      
      altCol := not altCol;
      
    end loop;
    
    altRow := not altRow;
    
  end loop;
  
  ExcelGen.setSheetProperties(ctx, sheet1, p_showGridLines => false);
  ExcelGen.setSheetProperties(ctx, sheet1, p_showRowColHeaders => false);
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'weave-pattern.xlsx');
  ExcelGen.closeContext(ctx);
  
end;
/
