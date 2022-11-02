declare
  sheet1  ExcelGen.sheetHandle;
  table1  ExcelGen.tableHandle;
  ctx     ExcelGen.ctxHandle;
  rc      sys_refcursor;
begin

  ctx := ExcelGen.createContext(ExcelGen.FILE_XLSX);
  
  sheet1 := ExcelGen.addSheet(ctx, 'test');
  ExcelGen.setDateFormat(ctx, sheet1, 'dd/mm/yyyy');
  ExcelGen.setTimestampFormat(ctx, 'dd/mm/yyyy hh:mm:ss.000');

  open rc for
  select anydata.convertnumber(123) c1 from dual union all
  select anydata.convertvarchar2('ABC') from dual union all
  select anydata.convertdate(sysdate) from dual union all
  select anydata.converttimestamptz(systimestamp) from dual
  ;

  table1 := ExcelGen.addTable(ctx, sheet1, rc);
  ExcelGen.setColumnProperties(ctx, sheet1, 1, p_width => 22);
  
  ExcelGen.createFile(ctx, 'TEST_DIR', 'anydata-column.xlsx');
  ExcelGen.closeContext(ctx);

end;
/
