alter session set plsql_optimize_level=3;
set define on

prompt set no_crypto to TRUE if you do not want to install xutl_offcrpto which requires DBMS_CRYPTO
--define no_crypto=TRUE
define no_crypto=FALSE

-- for conditional compilation based on sqlplus define settings.
-- When we select a column alias named "file_choice", we get a sqlplus define value for "file_choice"
COLUMN file_choice NEW_VALUE do_file NOPRINT

prompt Creating package XUTL_CDF ...
@@MSUtilities/CDFManager/xutl_cdf.pks
@@MSUtilities/CDFManager/xutl_cdf.pkb

SELECT DECODE('&&no_crypto','TRUE','do_nothing.sql xutl_offcrypto.pks', 'MSUtilities/OfficeCrypto/xutl_offcrypto.pks') AS file_choice FROM dual;
prompt calling &&do_file
@@&&do_file
--@@MSUtilities/OfficeCrypto/xutl_offcrypto.pks
SELECT DECODE('&&no_crypto','TRUE','do_nothing.sql xutl_offcrypto.pkb', 'MSUtilities/OfficeCrypto/xutl_offcrypto.pkb') AS file_choice FROM dual;
prompt calling &&do_file
@@&&do_file
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
ALTER SESSION SET plsql_ccflags='no_crypto:&&no_crypto';
@@plsql/ExcelGen.pks
@@plsql/ExcelGen.pkb
ALTER SESSION SET plsql_ccflags='';
