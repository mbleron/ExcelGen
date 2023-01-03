declare

  ctx     ExcelGen.ctxHandle;
  sheet1  ExcelGen.sheetHandle;
  sheet2  ExcelGen.sheetHandle;
  fill    ExcelGen.CT_Fill;

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
  
  ctx := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  
  -- sheet1
  sheet1 := ExcelGen.addSheet(ctx, 'sheet1');
  
  fill := ExcelGen.makeGradientFill();
  
  for i in 0 .. 9 loop
    ExcelGen.addGradientStop(fill, i/9, hsl2rgb(i*40,1,.5));
  end loop;
  
  ExcelGen.putCell(ctx, sheet1, 1, 1, p_style => ExcelGen.makeCellStyle(ctx, p_fill => fill));
  ExcelGen.setColumnProperties(ctx, sheet1, 1, p_width => 200);
  ExcelGen.setRowProperties(ctx, sheet1, 1, p_height => 30);
  
  -- sheet2
  sheet2 := ExcelGen.addSheet(ctx, 'sheet2');
  
  ExcelGen.putCell(ctx, sheet2, 2, 3, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to right,#FF0000,#FFAA00,#AAFF00,#00FF00,#00FFAA,#00AAFF,#0000FF,#AA00FF,#FF00AA)'));
  ExcelGen.putCell(ctx, sheet2, 2, 4, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to top right,#FF00AA 50%,#FFFFFF 50%)'));
  ExcelGen.putCell(ctx, sheet2, 3, 4, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to top,#FF0000,#FFAA00,#AAFF00,#00FF00,#00FFAA,#00AAFF,#0000FF,#AA00FF,#FF00AA)'));
  ExcelGen.putCell(ctx, sheet2, 4, 4, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to bottom right,#FF0000 50%,#FFFFFF 50%)'));
  ExcelGen.putCell(ctx, sheet2, 4, 3, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to left,#FF0000,#FFAA00,#AAFF00,#00FF00,#00FFAA,#00AAFF,#0000FF,#AA00FF,#FF00AA)'));
  ExcelGen.putCell(ctx, sheet2, 4, 2, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to bottom left,#FF00AA 50%,#FFFFFF 50%)'));
  ExcelGen.putCell(ctx, sheet2, 3, 2, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to bottom,#FF0000,#FFAA00,#AAFF00,#00FF00,#00FFAA,#00AAFF,#0000FF,#AA00FF,#FF00AA)'));
  ExcelGen.putCell(ctx, sheet2, 2, 2, p_style => ExcelGen.makeCellStyleCss(ctx, 'background:linear-gradient(to top left,#FF0000 50%,#FFFFFF 50%)'));
  
  ExcelGen.setColumnProperties(ctx, sheet2, 1, p_width => 5);
  ExcelGen.setColumnProperties(ctx, sheet2, 2, p_width => 5);
  ExcelGen.setColumnProperties(ctx, sheet2, 3, p_width => 40);
  ExcelGen.setColumnProperties(ctx, sheet2, 4, p_width => 5);
  ExcelGen.setRowProperties(ctx, sheet2, 1, p_height => 30);
  ExcelGen.setRowProperties(ctx, sheet2, 2, p_height => 30);
  ExcelGen.setRowProperties(ctx, sheet2, 3, p_height => 213.75);
  ExcelGen.setRowProperties(ctx, sheet2, 4, p_height => 30);
  
  ExcelGen.setSheetProperties(ctx, sheet2, p_showGridLines => false);
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'rainbow-gradient.xlsx');
  ExcelGen.closeContext(ctx);
  
end;
/
