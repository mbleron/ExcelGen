declare

  ctx     ExcelGen.ctxHandle := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  sheet1  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'sheet1');
  sheet2  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'sheet2');
  sheet3  ExcelGen.sheetHandle := ExcelGen.addSheet(ctx, 'sheet3', p_state => ExcelGen.ST_VERYHIDDEN);
  table1  ExcelGen.tableHandle := ExcelGen.addTable(ctx, sheet2, 'select empno, ename, job, sal, comm from scott.emp');
  table2  ExcelGen.tableHandle := ExcelGen.addTable(ctx, sheet3, 'select distinct job from scott.emp');
  
  style1  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'border:solid medium black');
  style2  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background-color:#C1F0C8');
  style3  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'background-color:#FBE2D5');
  style4  ExcelGen.cellStyleHandle := ExcelGen.makeCellStyleCss(ctx, 'font-weight:bold');
  
begin
  
  ExcelGen.putStringCell(ctx, sheet1, 1, 1, 'Use tolerance?');
  ExcelGen.putStringCell(ctx, sheet1, 2, 1, 'Day tolerance');
  ExcelGen.putStringCell(ctx, sheet1, 3, 1, 'Input date');
  ExcelGen.putStringCell(ctx, sheet1, 5, 1, 'Minimum');
  ExcelGen.putStringCell(ctx, sheet1, 6, 1, 'Maximum');
  ExcelGen.putStringCell(ctx, sheet1, 7, 1, 'Start date');
  ExcelGen.putStringCell(ctx, sheet1, 8, 1, 'End date');
  
  ExcelGen.putStringCell(ctx, sheet1, 1, 2, 'Yes');
  ExcelGen.putNumberCell(ctx, sheet1, 2, 2, 1);
  ExcelGen.putDateCell(ctx, sheet1, 3, 2, trunc(sysdate));
  ExcelGen.putNumberCell(ctx, sheet1, 5, 2, 1);
  ExcelGen.putNumberCell(ctx, sheet1, 6, 2, 10);
  ExcelGen.putDateCell(ctx, sheet1, 7, 2, date '2024-01-01');
  ExcelGen.putDateCell(ctx, sheet1, 8, 2, date '2024-12-31');
  
  ExcelGen.setDateFormat(ctx, sheet1, 'dd/mm/yyyy');
  ExcelGen.setColumnProperties(ctx, sheet1, 1, p_width => 14);
  
  ExcelGen.setRangeStyle(ctx, sheet1, 'A1:B3', style1);
  ExcelGen.setRangeStyle(ctx, sheet1, 'A5:B8', style1);
  ExcelGen.setRangeStyle(ctx, sheet1, 'A1:A3', style4);
  ExcelGen.setRangeStyle(ctx, sheet1, 'A5:A8', style4);
  ExcelGen.setRangeStyle(ctx, sheet1, 'B1:B3', style2);
  ExcelGen.setRangeStyle(ctx, sheet1, 'B5:B8', style3);
  
  ExcelGen.addDataValidationRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => 'list'
  , p_cellRange => ExcelTypes.ST_Sqref('B1')
  , p_value1    => '"Yes,No"'
  );

  ExcelGen.addDataValidationRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => 'whole'
  , p_cellRange => ExcelTypes.ST_Sqref('B2')
  , p_operator  => 'between'
  , p_value1    => '$B$5'
  , p_value2    => '$B$6'
  , p_showInputMessage => true
  , p_promptMsg => 'Enter a value between Minimum & Maximum' 
  );

  ExcelGen.addDataValidationRule(
    p_ctxId     => ctx
  , p_sheetId   => sheet1
  , p_type      => 'date'
  , p_cellRange => ExcelTypes.ST_Sqref('B3')
  , p_operator  => 'between'
  , p_value1    => '$B$7-IF($B$1="Yes",$B$2,0)'
  , p_value2    => '$B$8+IF($B$1="Yes",$B$2,0)'
  , p_showErrorMessage => true
  , p_errorTitle => 'Input date'
  , p_errorMsg   => 'Input date out of range.'
  );

  ExcelGen.setTableHeader(ctx, sheet2, table1);
  ExcelGen.setTableProperties(ctx, sheet2, table1, p_style => 'TableStyleLight1');
  ExcelGen.setTableProperties(ctx, sheet3, table2, p_tableName => 'JobList');
  ExcelGen.setTableColumnValidationRule(ctx, sheet2, table1, 3, 'list', 'INDIRECT("JobList")');
  ExcelGen.setTableColumnValidationRule(ctx, sheet2, table1, 5, 'decimal', p_operator => 'lessThanOrEqual', p_value1 => 'D2', p_refStyle1 => ExcelFmla.REF_A1);
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'test-dataval.xlsx');
  ExcelGen.closeContext(ctx); 
end;
/