CREATE OR REPLACE FUNCTION test_num_fmt RETURN BLOB AS
    v_blob          BLOB;
    v_src           SYS_REFCURSOR;
    v_ctxId         ExcelGen.ctxHandle;
    v_sheetHandle   BINARY_INTEGER;
BEGIN
    OPEN v_src FOR
        SELECT TO_CHAR(e.employee_id) AS employee_id, e.last_name, e.first_name, d.department_name, e.salary
        FROM hr.employees e
        INNER JOIN hr.departments d
            ON d.department_id = e.department_id
        UNION ALL
        SELECT '999' AS employee_id, 'Baggins' As last_name, 'Bilbo' as first_name, 'Sales' AS department_name
            ,123.45 AS salary
        FROM dual
        ;

        v_ctxId := ExcelGen.createContext();
        v_sheetHandle := ExcelGen.addSheetFromCursor(v_ctxId, 'Employee Salaries', v_src, p_sheetIndex => 1);
        BEGIN
            CLOSE v_src;
        EXCEPTION WHEN invalid_cursor THEN NULL;
        END;
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- style with alternating colors on each row. 
        ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');

        -- general settings for the entire workbook
        ExcelGen.setDateFormat(v_ctxId, 'mm/dd/yyyy');
        ExcelGen.setNumFormat(v_ctxId, '$0.00');

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;

END test_num_fmt;
/

CREATE OR REPLACE FUNCTION test_default RETURN BLOB AS
    v_blob          BLOB;
    v_src           SYS_REFCURSOR;
    v_ctxId         ExcelGen.ctxHandle;
    v_sheetHandle   BINARY_INTEGER;
BEGIN
    OPEN v_src FOR
        SELECT TO_CHAR(e.employee_id) AS employee_id, e.last_name, e.first_name, d.department_name, e.salary
        FROM hr.employees e
        INNER JOIN hr.departments d
            ON d.department_id = e.department_id
        UNION ALL
        SELECT '999' AS employee_id, 'Baggins' As last_name, 'Bilbo' as first_name, 'Sales' AS department_name
            ,123.45 AS salary
        FROM dual
        ;

        v_ctxId := ExcelGen.createContext();
        v_sheetHandle := ExcelGen.addSheetFromCursor(v_ctxId, 'Employee Salaries', v_src, p_sheetIndex => 1);
        BEGIN
            CLOSE v_src;
        EXCEPTION WHEN invalid_cursor THEN NULL;
        END;
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- style with alternating colors on each row. 
        ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');

        -- general settings for the entire workbook
        --ExcelGen.setDateFormat(v_ctxId, 'mm/dd/yyyy');
        --ExcelGen.setNumFormat(v_ctxId, '$0.00');

        v_blob := ExcelGen.getFileContent(v_ctxId);
        ExcelGen.closeContext(v_ctxId);
        RETURN v_blob;

END test_default;
/

