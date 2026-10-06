-- ====================================================================
-- SCRIPT FOR COMPANY DATABASE ADMINISTRATOR
-- Purpose: Create or replace the unified View V_EMP_DETAILS
-- Using the 3 source tables:
--   1. Temp_Sh_Lpt_Sep              (Columns: SLNO, GPFPRAN)
--   2. V_GpfpPran_Max                (Columns: ACCNO, PCNO)
--      [Note: ACCNO in this table refers to GPFPRAN]
--   3. V_Emp_Pis_BankAccountDetails  (Columns: PIS, PCNO, ACCNO)
--      [Note: ACCNO in this table is the Bank Account Number, NOT GPFPRAN]
--
-- Resulting Unified View:
--   V_EMP_DETAILS (Columns: SLNO, GPFPRAN, PCNO, PIS, BANK_ACCNO, ACCNO)
-- ====================================================================

-- OPTION A (RECOMMENDED): FULL OUTER JOIN
-- Ensures lookups by any key (GPFPRAN, PCNO, PIS, BANK_ACCNO/ACCNO) succeed,
-- linking all 3 tables seamlessly:
-- - Temp_Sh_Lpt_Sep.GPFPRAN matches V_GpfpPran_Max.ACCNO (both refer to GPFPRAN)
-- - V_GpfpPran_Max.PCNO matches V_Emp_Pis_BankAccountDetails.PCNO
-- - V_Emp_Pis_BankAccountDetails.ACCNO is the employee's Bank Account Number,
--   exposed as BANK_ACCNO (and ACCNO alias) so both queries succeed.

CREATE OR REPLACE VIEW V_EMP_DETAILS AS
SELECT 
    a.SLNO,
    COALESCE(a.GPFPRAN, b.ACCNO) AS GPFPRAN,
    COALESCE(b.PCNO, c.PCNO) AS PCNO,
    c.PIS,
    c.ACCNO AS BANK_ACCNO,
    c.ACCNO AS ACCNO
FROM Temp_Sh_Lpt_Sep a
FULL OUTER JOIN V_GpfpPran_Max b ON a.GPFPRAN = b.ACCNO
FULL OUTER JOIN V_Emp_Pis_BankAccountDetails c ON b.PCNO = c.PCNO;

-- --------------------------------------------------------------------
-- OPTION B: Standard LEFT JOIN (If workflow always starts from Temp_Sh_Lpt_Sep):
-- --------------------------------------------------------------------
-- CREATE OR REPLACE VIEW V_EMP_DETAILS AS
-- SELECT 
--     a.SLNO,
--     a.GPFPRAN,
--     b.PCNO,
--     c.PIS,
--     c.ACCNO AS BANK_ACCNO,
--     c.ACCNO AS ACCNO
-- FROM Temp_Sh_Lpt_Sep a
-- LEFT JOIN V_GpfpPran_Max b ON a.GPFPRAN = b.ACCNO
-- LEFT JOIN V_Emp_Pis_BankAccountDetails c ON b.PCNO = c.PCNO;

-- Permissions:
-- GRANT SELECT ON V_EMP_DETAILS TO <your_app_db_user>;
