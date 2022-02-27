--set serveroutput on
DECLARE
    c       SYS_REFCURSOR;
    c2      SYS_REFCURSOR;
    v_blob  BLOB;
    v_sql   CLOB := q'!WITH add_bilbo AS (
            SELECT e.employee_id AS employee_id, e.last_name, e.first_name, d.department_name, e.salary
            FROM hr.employees e
            INNER JOIN hr.departments d
                ON d.department_id = e.department_id
            UNION ALL
            SELECT 999 AS employee_id, 'Baggins' As last_name, 'Bilbo' as first_name, 'Sales' AS department_name
                ,123.45 AS salary
            FROM dual
        ) SELECT employee_id, last_name, first_name, department_name
                ,TO_BINARY_DOUBLE(salary) AS salary
        FROM add_bilbo ORDER BY last_name, first_name!';

    FUNCTION get_xlsx(p_src SYS_REFCURSOR) 
    RETURN BLOB AS
        v_blob          BLOB;
        v_ctxId         ExcelGen.ctxHandle;
        v_sheetHandle   BINARY_INTEGER;
    BEGIN
        v_ctxId := ExcelGen.createContext();
        ExcelGen.setDebug(TRUE);
        v_sheetHandle := ExcelGen.addSheetFromCursor(v_ctxId, 'Employee Salaries', p_src, p_sheetIndex => 1);
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- single column format on the salary column. The ID column keeps default format
        ExcelGen.setColumnFormat(
            p_ctxId     => v_ctxId
            ,p_sheetId  => v_sheetHandle
            ,p_columnId => 5        -- the salary column
            ,p_format   => '$#,##0.00'
        );
        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'Employee Salaries2', v_sql, p_sheetIndex => 2);

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;
    END;

    FUNCTION get_xlsb(p_src SYS_REFCURSOR) 
    RETURN BLOB AS
        v_blob          BLOB;
        v_ctxId         ExcelGen.ctxHandle;
        v_sheetHandle   BINARY_INTEGER;
    BEGIN
        v_ctxId := ExcelGen.createContext(ExcelGen.FILE_XLSB);
        ExcelGen.setDebug(TRUE);
        v_sheetHandle := ExcelGen.addSheetFromCursor(v_ctxId, 'Employee Salaries', p_src, p_sheetIndex => 1);
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- single column format on the salary column. The ID column keeps default format
        ExcelGen.setColumnFormat(
            p_ctxId     => v_ctxId
            ,p_sheetId  => v_sheetHandle
            ,p_columnId => 5        -- the salary column
            ,p_format   => '$#,##0.00'
        );
        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'Employee Salaries2', v_sql, p_sheetIndex => 2);

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;
    END;
BEGIN

    OPEN c FOR v_sql ;
    v_blob := get_xlsx(c);
    OPEN c2 FOR v_sql ;
    v_blob := get_xlsb(c2);
END;
/

