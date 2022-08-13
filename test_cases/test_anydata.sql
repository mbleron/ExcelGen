whenever sqlerror continue
drop table test_anydata;
whenever sqlerror exit failure
  CREATE TABLE "TEST_ANYDATA" 
   (	"DATA_ROW_NR" NUMBER, 
	"id" "ANYDATA" , 
	"data" "ANYDATA" , 
	"ddata" "ANYDATA" 
   ) ;
REM INSERTING into TEST_ANYDATA
Insert into TEST_ANYDATA (DATA_ROW_NR,"id","data","ddata") values (1
    ,SYS.ANYDATA.convertVarchar2('varchar')
    ,SYS.ANYDATA.convertVarchar2('test data')
    ,null
);
Insert into TEST_ANYDATA (DATA_ROW_NR,"id","data","ddata") values (2
    ,SYS.ANYDATA.convertVarchar2('number')
    ,SYS.ANYDATA.convertNumber(123.27)
    ,null
);
Insert into TEST_ANYDATA (DATA_ROW_NR,"id","data","ddata") values (3
    ,SYS.ANYDATA.convertVarchar2('date')
    ,SYS.ANYDATA.convertDate(TO_DATE('08/02/2022','mm/dd/yyyy'))
    ,SYS.ANYDATA.convertDate(TO_DATE('08/21/2022','mm/dd/yyyy'))
);
Insert into TEST_ANYDATA (DATA_ROW_NR,"id","data","ddata") values (4
    ,SYS.ANYDATA.convertVarchar2('varcharmore')
    ,SYS.ANYDATA.convertVarchar2('more test data')
    ,SYS.ANYDATA.convertVarchar2('text')
);
COMMIT;
--
WITH 
FUNCTION get_xlsx(p_src SYS_REFCURSOR) RETURN BLOB AS
    v_blob          BLOB;
    v_ctxId         ExcelGen.ctxHandle;
    v_sheetHandle   BINARY_INTEGER;
BEGIN
    -- pick one or the other of these to test binfile vs xlsx
        --v_ctxId := ExcelGen.createContext(ExcelGen.FILE_XLSX);
        v_ctxId := ExcelGen.createContext(ExcelGen.FILE_XLSB);
        v_sheetHandle := ExcelGen.addSheetFromCursor(v_ctxId, 'test anydata', p_src, p_sheetIndex => 1);
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- style with alternating colors on each row. 
        ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');
        -- single column format is necessary for any column that is primarily dates and recommended
        -- for any column that is primarily number
        ExcelGen.setColumnFormat(
            p_ctxId     => v_ctxId
            ,p_sheetId  => v_sheetHandle
            ,p_columnId => 4       
            ,p_format   => 'mm/dd/yyyy' --'$#,##0.00'
        );
        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;
END;
a AS (
    SELECT * FROM test_anydata
) SELECT get_xlsx(CURSOR(SELECT * FROM a)) FROM dual
;
/

