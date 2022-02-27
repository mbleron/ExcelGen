--set serveroutput on
DECLARE
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

    FUNCTION get_xlsx
    RETURN BLOB AS
        v_blob          BLOB;
        v_ctxId         ExcelGen.ctxHandle;
        v_sheetHandle   BINARY_INTEGER;
    BEGIN
        v_ctxId := ExcelGen.createContext();
        ExcelGen.setDebug(TRUE);
        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'Employee Salaries2', v_sql, p_sheetIndex => 1);
        ExcelGen.setBindVariable(v_ctxId, v_sheetHandle, 'bogusbind', 'bogusval');

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;
    END;

    FUNCTION get_xlsb
    RETURN BLOB AS
        v_blob          BLOB;
        v_ctxId         ExcelGen.ctxHandle;
        v_sheetHandle   BINARY_INTEGER;
    BEGIN
        v_ctxId := ExcelGen.createContext(ExcelGen.FILE_XLSB);
        ExcelGen.setDebug(TRUE);
        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'Employee Salaries2', v_sql, p_sheetIndex => 1);
        ExcelGen.setBindVariable(v_ctxId, v_sheetHandle, 'bogusbind', 'bogusval');

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;
    END;
BEGIN

    BEGIN
    v_blob := get_xlsx;
    EXCEPTION WHEN OTHERS THEN NULL;
    END;
    v_blob := get_xlsb;
END;
/

