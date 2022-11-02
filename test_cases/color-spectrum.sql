declare

  sheet1  ExcelGen.sheetHandle;
  ctx     ExcelGen.ctxHandle;
  rowIdx  pls_integer := 0;
  colIdx  pls_integer := 0;
  
  function hsl2rgb2 (H in number, SL in number, L in number)
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

  ctx := ExcelGen.createContext();
  sheet1 := ExcelGen.addSheet(ctx, 'sheet1');
  ExcelGen.setSheetProperties(ctx, sheet1, p_showGridLines => false, p_showRowColHeaders => false);

  for lt in 0 .. 100 loop
    rowIdx := rowIdx + 1;
    colIdx := 0;
    for v in 0 .. 359 loop
      colIdx := colIdx + 1;
      ExcelGen.putCell(ctx, sheet1, rowIdx, colIdx, null, p_style => ExcelGen.makeCellStyle(ctx, p_fill => ExcelGen.makePatternFill('solid',hsl2rgb2(v, 1, lt/100))));
    end loop;
  end loop;
  
  for colIdx in 1 .. 360 loop
    ExcelGen.setColumnProperties(ctx, sheet1, colIdx, p_width => 0.5);
  end loop;

  for rowIdx in 1 .. 101 loop
    ExcelGen.setRowProperties(ctx, sheet1, rowIdx, p_height => 6.75);
  end loop;

  ExcelGen.createFile(ctx, 'TEST_DIR', 'color-spectrum.xlsx');
  ExcelGen.closeContext(ctx);

end;
/
