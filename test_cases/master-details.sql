declare
  ctx       ExcelGen.ctxHandle;
  sheet1    ExcelGen.sheetHandle;
  tableId   ExcelGen.tableHandle;
  rowIdx    pls_integer := 0;
  empQuery  varchar2(1000) := 'select first_name, last_name, hire_date, job_id, salary from hr.employees where department_id = :1';
begin
  ctx := ExcelGen.createContext();
  
  sheet1 := ExcelGen.addSheet(ctx, 'Departments');
  
  for d in ( select department_id, department_name from hr.departments ) loop
    if tableId is null then
      tableId := ExcelGen.addTable(ctx, sheet1, empQuery, p_anchorRowOffset => 2, p_anchorColOffset => 1);
    else
      tableId := ExcelGen.addTable(ctx, sheet1, empQuery, p_anchorRowOffset => 3, p_anchorColOffset => 0, p_anchorTableId => tableId, p_anchorPosition => ExcelGen.BOTTOM_LEFT);
    end if;
    ExcelGen.setBindVariable(ctx, sheet1, tableId, '1', d.department_id);
    ExcelGen.putStringCell(ctx, sheet1, -1, 0, d.department_name, p_style => ExcelGen.makeCellStyleCss(ctx, 'font-weight:bold;background:yellowgreen;'), p_anchorTableId => tableId, p_anchorPosition => ExcelGen.TOP_LEFT);
    ExcelGen.mergeCells(ctx, sheet1, -1, 0, 1, 5, tableId, ExcelGen.TOP_LEFT);
  end loop;
  
  ExcelGen.setDateFormat(ctx, 'dd/mm/yyyy');
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'master-details.xlsx');
  ExcelGen.closeContext(ctx);
end;
/
