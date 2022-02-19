-- avoids compiling procedures that require dbms_crypto. Use this install script if dbms_crypto was
-- not granted to your schema
--
alter session set plsql_optimize_level=3;

prompt Creating package XUTL_CDF ...
@@MSUtilities/CDFManager/xutl_cdf.pks
@@MSUtilities/CDFManager/xutl_cdf.pkb

--@@MSUtilities/OfficeCrypto/xutl_offcrypto.pks
--@@MSUtilities/OfficeCrypto/xutl_offcrypto.pkb

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
-- this compile directive prevents compiling a procedure that requires dbms_crpto
ALTER SESSION SET plsql_ccflags='no_crypto:TRUE';
@@plsql/ExcelGen.pks
@@plsql/ExcelGen.pkb
ALTER SESSION SET plsql_ccflags='';
