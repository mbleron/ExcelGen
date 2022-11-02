declare

  ctx      ExcelGen.ctxHandle;
  sheet1   ExcelGen.sheetHandle;
  table1   ExcelGen.tableHandle;
  table2   ExcelGen.tableHandle;
  table3   ExcelGen.tableHandle;
  
begin

  ctx := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  sheet1 := ExcelGen.addSheet(ctx, 'sheet1');
  
  table1 := ExcelGen.addTable(ctx, sheet1, 'select employee_id, first_name, last_name from hr.employees where department_id = 50');
  table2 := ExcelGen.addTable(
              ctx
            , sheet1
            , 'select employee_id, first_name, last_name from hr.employees where department_id = 30'
            , p_anchorRowOffset => -1
            , p_anchorColOffset => 2
            , p_anchorTableId => table1
            , p_anchorPosition => ExcelGen.BOTTOM_RIGHT
            );
            
  table3 := ExcelGen.addTable(
              ctx
            , sheet1
            , 'select employee_id, first_name, last_name, hire_date from hr.employees where department_id = 60'
            , p_anchorRowOffset => 2 
            , p_anchorColOffset => -1
            , p_anchorTableId => table2
            , p_anchorPosition => ExcelGen.BOTTOM_LEFT
            );
            
  ExcelGen.putStringCell(ctx, sheet1, -2, 0, 'TEST1', p_anchorTableId => table2, p_anchorPosition => ExcelGen.TOP_LEFT);
  ExcelGen.putStringCell(ctx, sheet1, 1, 0, 'TEST2', p_anchorTableId => table3, p_anchorPosition => ExcelGen.BOTTOM_RIGHT);
            
  ExcelGen.setTableHeader(ctx, sheet1, table1, ExcelGen.makeCellStyleCss(ctx, 'background:lightgray'));
  ExcelGen.setTableHeader(ctx, sheet1, table2, ExcelGen.makeCellStyleCss(ctx, 'background:yellowgreen'));
  ExcelGen.setTableHeader(ctx, sheet1, table3, ExcelGen.makeCellStyleCss(ctx, 'background:yellowgreen'));
  
  ExcelGen.setSheetProperties(ctx, sheet1, 'A2');
  ExcelGen.setColumnProperties(ctx, sheet1, 7, p_width => 18);
  
  ExcelGen.setTableRowProperties(ctx, sheet1, table2, 3, ExcelGen.makeCellStyleCss(ctx, 'color:red;font-weight:bold'));
  ExcelGen.setTableProperties(ctx, sheet1, table2, 'tableStyleLight1');
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'relative-positioning.xlsx');
  ExcelGen.closeContext(ctx);
  
end;
/
