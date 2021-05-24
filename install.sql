alter session set plsql_optimize_level=3;

prompt Creating type ExcelTableCell ...
@@ExcelCommons/plsql/ExcelTableCell.tps

prompt Creating type ExcelTableCellList ...
@@ExcelCommons/plsql/ExcelTableCellList.tps

prompt Creating package ExcelTypes ...
@@ExcelCommons/plsql/ExcelTypes.pks
@@ExcelCommons/plsql/ExcelTypes.pkb

prompt Creating package XUTL_XLSB ...
@@ExcelCommons/plsql/xutl_xlsb.pks
@@ExcelCommons/plsql/xutl_xlsb.pkb

prompt Creating package ExcelGen ...
@@plsql/ExcelGen.pks
@@plsql/ExcelGen.pkb
