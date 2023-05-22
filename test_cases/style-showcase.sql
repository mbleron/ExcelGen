-- STYLE-SHOWCASE
declare

  ctx  ExcelGen.ctxHandle;

  -- Borders
  procedure makeBorderSheet is
    
    sheet1  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Borders');
    
  begin

    ExcelGen.setSheetProperties(ctx, sheet1, p_showGridLines => false);
    ExcelGen.setColumnProperties(ctx, sheet1, 1, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet1, 3, p_width => 1);
    ExcelGen.setColumnProperties(ctx, sheet1, 7, p_width => 1);

    ExcelGen.putCell(ctx, sheet1, 2, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:none'));
    ExcelGen.putStringCell(ctx, sheet1, 2, 4, 'none');

    ExcelGen.putCell(ctx, sheet1, 4, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:hairline'));
    ExcelGen.putStringCell(ctx, sheet1, 4, 4, 'hair');

    ExcelGen.putCell(ctx, sheet1, 6, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:dotted'));
    ExcelGen.putStringCell(ctx, sheet1, 6, 4, 'dotted');

    ExcelGen.putCell(ctx, sheet1, 8, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:thin dot-dot-dash'));
    ExcelGen.putStringCell(ctx, sheet1, 8, 4, 'dashDotDot');

    ExcelGen.putCell(ctx, sheet1, 10, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:thin dot-dash'));
    ExcelGen.putStringCell(ctx, sheet1, 10, 4, 'dashDot');

    ExcelGen.putCell(ctx, sheet1, 12, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:thin dashed'));
    ExcelGen.putStringCell(ctx, sheet1, 12, 4, 'dashed');

    ExcelGen.putCell(ctx, sheet1, 14, 2, null, ExcelGen.makeCellStyleCss(ctx, 'border:thin solid'));
    ExcelGen.putStringCell(ctx, sheet1, 14, 4, 'thin');


    ExcelGen.putCell(ctx, sheet1, 2, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:medium dot-dot-dash'));
    ExcelGen.putStringCell(ctx, sheet1, 2, 8, 'mediumDashDotDot');

    ExcelGen.putCell(ctx, sheet1, 4, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:dot-dash-slanted'));
    ExcelGen.putStringCell(ctx, sheet1, 4, 8, 'slantDashDot');

    ExcelGen.putCell(ctx, sheet1, 6, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:medium dot-dash'));
    ExcelGen.putStringCell(ctx, sheet1, 6, 8, 'mediumDashDot');

    ExcelGen.putCell(ctx, sheet1, 8, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:medium dashed'));
    ExcelGen.putStringCell(ctx, sheet1, 8, 8, 'mediumDashed');

    ExcelGen.putCell(ctx, sheet1, 10, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:medium solid'));
    ExcelGen.putStringCell(ctx, sheet1, 10, 8, 'medium');

    ExcelGen.putCell(ctx, sheet1, 12, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:thick solid'));
    ExcelGen.putStringCell(ctx, sheet1, 12, 8, 'thick');

    ExcelGen.putCell(ctx, sheet1, 14, 6, null, ExcelGen.makeCellStyleCss(ctx, 'border:double'));
    ExcelGen.putStringCell(ctx, sheet1, 14, 8, 'double');
    
  end;

  -- Patterns
  procedure makePatternSheet is
    
    sheet2  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Patterns');
    
  begin
    
    ExcelGen.setSheetProperties(ctx, sheet2, p_showGridLines => false);
    ExcelGen.setColumnProperties(ctx, sheet2, 1, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet2, 3, p_width => 1);
    ExcelGen.setColumnProperties(ctx, sheet2, 7, p_width => 1);
    ExcelGen.setColumnProperties(ctx, sheet2, 11, p_width => 1);

    ExcelGen.putCell(ctx, sheet2, 2, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:none;background:black'));
    ExcelGen.putStringCell(ctx, sheet2, 2, 4, 'solid');

    ExcelGen.putCell(ctx, sheet2, 4, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:gray-75'));
    ExcelGen.putStringCell(ctx, sheet2, 4, 4, 'darkGray');

    ExcelGen.putCell(ctx, sheet2, 6, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:gray-50'));
    ExcelGen.putStringCell(ctx, sheet2, 6, 4, 'mediumGray');

    ExcelGen.putCell(ctx, sheet2, 8, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:gray-25'));
    ExcelGen.putStringCell(ctx, sheet2, 8, 4, 'lightGray');

    ExcelGen.putCell(ctx, sheet2, 10, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:gray-125'));
    ExcelGen.putStringCell(ctx, sheet2, 10, 4, 'gray125');

    ExcelGen.putCell(ctx, sheet2, 12, 2, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:gray-0625'));
    ExcelGen.putStringCell(ctx, sheet2, 12, 4, 'gray0625');


    ExcelGen.putCell(ctx, sheet2, 2, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:horz-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 2, 8, 'darkHorizontal');

    ExcelGen.putCell(ctx, sheet2, 4, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:vert-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 4, 8, 'darkVertical');

    ExcelGen.putCell(ctx, sheet2, 6, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:reverse-dark-down'));
    ExcelGen.putStringCell(ctx, sheet2, 6, 8, 'darkDown');

    ExcelGen.putCell(ctx, sheet2, 8, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:diag-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 8, 8, 'darkUp');

    ExcelGen.putCell(ctx, sheet2, 10, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:diag-cross'));
    ExcelGen.putStringCell(ctx, sheet2, 10, 8, 'darkGrid');

    ExcelGen.putCell(ctx, sheet2, 12, 6, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thick-diag-cross'));
    ExcelGen.putStringCell(ctx, sheet2, 12, 8, 'darkTrellis');


    ExcelGen.putCell(ctx, sheet2, 2, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-horz-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 2, 12, 'lightHorizontal');

    ExcelGen.putCell(ctx, sheet2, 4, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-vert-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 4, 12, 'lightVertical');

    ExcelGen.putCell(ctx, sheet2, 6, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-reverse-diag-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 6, 12, 'lightDown');

    ExcelGen.putCell(ctx, sheet2, 8, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-diag-stripe'));
    ExcelGen.putStringCell(ctx, sheet2, 8, 12, 'lightUp');

    ExcelGen.putCell(ctx, sheet2, 10, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-horz-cross'));
    ExcelGen.putStringCell(ctx, sheet2, 10, 12, 'lightGrid');

    ExcelGen.putCell(ctx, sheet2, 12, 10, null, ExcelGen.makeCellStyleCss(ctx, 'mso-pattern:thin-diag-cross'));
    ExcelGen.putStringCell(ctx, sheet2, 12, 12, 'lightTrellis');
    
  end;

  -- Alignments
  procedure makeAlignmentSheet is
    
    rowIdx  pls_integer := 1;
    sheet3  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Alignments');

    colHeaderStyle  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background:#ffd966;text-align:center;vertical-align:middle;font-weight:bold;border-bottom:thin solid black');
    rowHeaderStyle  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background:#c6e0b4;text-align:center;vertical-align:middle;font-weight:bold;border-right:thin solid black');

    colHeaderStyle2  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'border-bottom:thin solid black');
    colHeaderStyle3  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background:#ffc000;text-align:right;vertical-align:middle;border-right:thin solid black');
    rowHeaderStyle2  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background:#70ad47;text-align:left;vertical-align:middle;border-right:thin solid black;border-bottom:thin solid black');

  begin

    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, to_char(unistr('Horizontal \2192')), colHeaderStyle3);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'left', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'center', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'right', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'fill', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'justify', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'distributed', colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'centerCountinuous', colHeaderStyle);

    ExcelGen.mergeCells(ctx, sheet3, 'B1:B2');
    ExcelGen.mergeCells(ctx, sheet3, 'C1:C2');
    ExcelGen.mergeCells(ctx, sheet3, 'D1:D2');
    ExcelGen.mergeCells(ctx, sheet3, 'E1:E2');
    ExcelGen.mergeCells(ctx, sheet3, 'F1:F2');
    ExcelGen.mergeCells(ctx, sheet3, 'G1:G2');
    ExcelGen.mergeCells(ctx, sheet3, 'H1:H2');

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, to_char(unistr('\2193 Vertical')), rowHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 2, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 3, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 4, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 5, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 6, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 7, null, colHeaderStyle2);
    ExcelGen.putCell(ctx, sheet3, rowIdx, 8, null, colHeaderStyle2);

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, 'top', rowHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'X');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'XXX XXX XXX');

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, 'center', rowHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'X');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'XXX XXX XXX');

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, 'bottom', rowHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'X');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'XXX XXX XXX');

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, 'justify', rowHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'X');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'XXX XXX XXX XXX');

    rowIdx := rowIdx + 1;
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 1, 'distributed', rowHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 2, 'XXX XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 3, 'XXX XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 4, 'XXX XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 5, 'X');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 6, 'XXX XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 7, 'XXX XXX XXX XXX XXX');
    ExcelGen.putStringCell(ctx, sheet3, rowIdx, 8, 'XXX XXX XXX XXX XXX');

    ExcelGen.setColumnProperties(ctx, sheet3, 1, p_width => 16);
    ExcelGen.setColumnProperties(ctx, sheet3, 2, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:left'));
    ExcelGen.setColumnProperties(ctx, sheet3, 3, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:center'));
    ExcelGen.setColumnProperties(ctx, sheet3, 4, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:right'));
    ExcelGen.setColumnProperties(ctx, sheet3, 5, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:fill'));
    ExcelGen.setColumnProperties(ctx, sheet3, 6, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:justify'));
    ExcelGen.setColumnProperties(ctx, sheet3, 7, p_width => 11, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:distributed'));
    ExcelGen.setColumnProperties(ctx, sheet3, 8, p_width => 18, p_style => ExcelGen.makeCellStyleCss(ctx,'text-align:center-across'));

    ExcelGen.setRowProperties(ctx, sheet3, 3, p_height => 50, p_style => ExcelGen.makeCellStyleCss(ctx,'vertical-align:top'));
    ExcelGen.setRowProperties(ctx, sheet3, 4, p_height => 50, p_style => ExcelGen.makeCellStyleCss(ctx,'vertical-align:middle'));
    ExcelGen.setRowProperties(ctx, sheet3, 5, p_height => 50, p_style => ExcelGen.makeCellStyleCss(ctx,'vertical-align:bottom'));
    ExcelGen.setRowProperties(ctx, sheet3, 6, p_height => 50, p_style => ExcelGen.makeCellStyleCss(ctx,'vertical-align:justify'));
    ExcelGen.setRowProperties(ctx, sheet3, 7, p_height => 50, p_style => ExcelGen.makeCellStyleCss(ctx,'vertical-align:distributed'));

  end;

  -- Colors
  procedure makeColorSheet is
    
    rowIdx    pls_integer := 1;
    colorMap  ExcelTypes.colorMap_t := ExcelTypes.getColorMap();
    colorName varchar2(20);
    colorCode varchar2(6);

    colHeaderStyle  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background:#ffc000;text-align:center;vertical-align:middle;font-weight:bold');
    colorCodeStyle  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'font:10pt "Courier New"');

    sheet4  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Colors');
    
  begin

    ExcelGen.putStringCell(ctx, sheet4, rowIdx, 1, 'Name', p_style => colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet4, rowIdx, 2, 'Code', p_style => colHeaderStyle);
    ExcelGen.putStringCell(ctx, sheet4, rowIdx, 3, 'Sample', p_style => colHeaderStyle);

    ExcelGen.setColumnProperties(ctx, sheet4, 1, p_width => 20);
    ExcelGen.setSheetProperties(ctx, sheet4, p_activePaneAnchorRef => 'A2');

    colorName := colorMap.first;
    while colorName is not null loop
      rowIdx := rowIdx + 1;
      colorCode := colorMap(colorName);
      ExcelGen.putStringCell(ctx, sheet4, rowIdx, 1, colorName);
      ExcelGen.putStringCell(ctx, sheet4, rowIdx, 2, colorCode, colorCodeStyle);
      ExcelGen.putCell(ctx, sheet4, rowIdx, 3, null, ExcelGen.makeCellStyle(ctx, p_fill => ExcelGen.makePatternFill('solid','#'||colorCode)));
      colorName := colorMap.next(colorName);
    end loop;
    
  end;

  procedure makeTableSheet is
    
    sheet5  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Tables');
    tableId  ExcelGen.tableHandle;

    procedure putTable (shortName in varchar2, startRow in pls_integer, nrows in pls_integer, cellCount in pls_integer) is
    begin
      ExcelGen.putStringCell( ctx
                      , sheet5
                      , p_rowIdx => startRow
                      , p_colIdx => 2
                      , p_value => shortName
                      , p_style => ExcelGen.makeCellStyleCss(ctx, 'font-size:10pt;font-weight:bold')
                      );
      for i in 1 .. cellCount loop
        tableId := ExcelGen.addTable( ctx
                                    , sheet5
                                    , 'select cast(null as number) c1 from dual connect by level <= 3'
                                    , p_anchorRowOffset => startRow + 2 + 6* mod(i-1,nrows)
                                    , p_anchorColOffset => 2*(trunc((i-1)/nrows) + 1)
                                    );
        ExcelGen.setTableHeader(ctx, sheet5, tableId);
        ExcelGen.setTableProperties(ctx, sheet5, tableId, p_style => 'TableStyle'||shortName||to_char(i));
        ExcelGen.putStringCell( ctx
                              , sheet5
                              , p_rowIdx => startRow + 6 + 6* mod(i-1,nrows)
                              , p_colIdx => 2*(trunc((i-1)/nrows)+1)
                              , p_value => 'TableStyle'||shortName||to_char(i)
                              , p_style => ExcelGen.makeCellStyleCss(ctx, 'font-size:10pt;text-align:center')
                              );
      end loop;
      
    end;

  begin
    
    ExcelGen.setSheetProperties(ctx, sheet5, p_showGridLines => false, p_defaultRowHeight => 12.75);
    ExcelGen.setColumnProperties(ctx, sheet5, 1, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet5, 2, p_width => 17);
    ExcelGen.setColumnProperties(ctx, sheet5, 3, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet5, 4, p_width => 17);
    ExcelGen.setColumnProperties(ctx, sheet5, 5, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet5, 6, p_width => 17);
    ExcelGen.setColumnProperties(ctx, sheet5, 7, p_width => 3);
    ExcelGen.setColumnProperties(ctx, sheet5, 8, p_width => 17);

    putTable('Light', 2, 7, 21);
    putTable('Medium', 47, 7, 28);
    putTable('Dark', 92, 4, 11);

  end;
  
  procedure makeGradientSheet is
    sheet6  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Gradients');
  begin
    ExcelGen.putCell(ctx, sheet6, 1, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background-image: linear-gradient(to top, #dbdcd7 0%, #dddcd7 24%, #e2c9cc 30%, #e7627d 46%, #b8235a 59%, #801357 71%, #3d1635 84%, #1c1a27 100%);'));
    ExcelGen.putCell(ctx, sheet6, 2, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background-image: linear-gradient(to top, #3f51b1 0%, #5a55ae 13%, #7b5fac 25%, #8f6aae 38%, #a86aa4 50%, #cc6b8e 62%, #f18271 75%, #f3a469 87%, #f7c978 100%);'));
    ExcelGen.putCell(ctx, sheet6, 3, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(89.2deg, rgb(0, 0, 0) 10.4%, rgb(255, 0, 0) 37.1%, rgb(255, 216, 51) 64.3%, rgb(255, 255, 255) 90.5%);'));
    ExcelGen.putCell(ctx, sheet6, 4, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(110.7deg, rgb(255, 81, 47) 1.7%, rgb(255, 167, 47) 8.2%, rgb(218, 253, 1) 16.2%, rgb(98, 234, 20) 23.4%, rgb(69, 193, 42) 32.8%, rgb(7, 249, 149) 43.7%, rgb(6, 200, 217) 55.3%, rgb(18, 51, 233) 65.5%, rgb(122, 59, 202) 74.5%, rgb(231, 7, 249) 82.3%, rgb(202, 59, 163) 91.4%);'));
    ExcelGen.putCell(ctx, sheet6, 5, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(-20deg, #d558c8 0%, #24d292 100%);'));
    ExcelGen.putCell(ctx, sheet6, 6, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(102.2deg, rgb(109, 118, 150) 7.1%, rgb(89, 72, 79) 14.7%, rgb(69, 92, 79) 28%, rgb(204, 85, 67) 43.4%, rgb(237, 181, 121) 60.2%, rgb(219, 230, 175) 76.8%);'));
    ExcelGen.putCell(ctx, sheet6, 7, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(to top, #0fd850 0%, #f9f047 100%);'));
    ExcelGen.putCell(ctx, sheet6, 8, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(105.9deg, rgb(15, 209, 165) 3.8%, rgb(15, 157, 209) 20.8%, rgb(133, 13, 230) 51.9%, rgb(230, 13, 202) 73.1%, rgb(242, 180, 107) 94.1%);'));
    ExcelGen.putCell(ctx, sheet6, 9, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient( 109.6deg, rgb(33, 25, 180) 11.2%, rgb(253, 29, 29) 55.2%, rgb(252, 176, 69) 91.1%);'));
    ExcelGen.putCell(ctx, sheet6, 10, 1, p_style => ExcelGen.makeCellStyleCss(ctx, 'background: linear-gradient(99.6deg, rgb(112, 128, 152) 10.6%, rgb(242, 227, 234) 32.9%, rgb(234, 202, 213) 52.7%, rgb(220, 227, 239) 72.8%, rgb(185, 205, 227) 81.1%, rgb(154, 180, 212) 100%);'));
    
    ExcelGen.setColumnProperties(ctx, sheet6, 1, p_width => 50);
    for i in 1 .. 10 loop
      ExcelGen.setRowProperties(ctx, sheet6, i, p_height => 195);
    end loop;    
  end;

  procedure makeRichTextSheet is
    sheet7  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'Rich Text');
    style1  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'font-size:20pt;font-weight:bold;text-align:center;vertical-align:middle');
  begin
    ExcelGen.putRichTextCell(
      p_ctxId => ctx
    , p_sheetId => sheet7
    , p_rowIdx => 1
    , p_colIdx => 1
    , p_value => 
      '<span style="color:#FF0000">R</span>
       <span style="color:#FFFF00">A</span>
       <span style="color:#00FF00">I</span>
       <span style="color:#00FFFF">N</span>
       <span style="color:#0000FF">B</span>
       <span style="color:#FF00FF">O</span>
       <span style="color:#FF0000">W</span>'
    , p_style => style1
    );

    ExcelGen.putRichTextCell(
      p_ctxId   => ctx
    , p_sheetId => sheet7
    , p_rowIdx  => 2
    , p_colIdx  => 1
    , p_value   => 
      'The chemical formula of glucose is <span style="color:blue;font-weight:bold">'||regexp_replace('C6H12O6','(\d+)','<sub>\1</sub>')||'</span>'
    );
    
    ExcelGen.setColumnProperties(ctx, sheet7, 1, p_width => ExcelGen.colPxToCharWidth(128));
    ExcelGen.setRowProperties(ctx, sheet7, 1, p_height => 30);    
  end;

begin

  ctx := ExcelGen.createContext(ExcelGen.FILE_XLSX);

  makeBorderSheet;
  makePatternSheet;
  makeTableSheet;
  makeAlignmentSheet;
  makeColorSheet;
  makeGradientSheet;
  makeRichTextSheet;

  ExcelGen.createFile(ctx, 'TEST_DIR', 'style-showcase.xlsx');
  ExcelGen.closeContext(ctx);

end;
/
