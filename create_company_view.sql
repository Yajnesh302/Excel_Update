-- ====================================================================
-- SCRIPT FOR COMPANY DATABASE ADMINISTRATOR
-- Purpose: Create or replace the unified View V_EMP_DETAILS
--
-- Background & Fix Summary:
--   1. Temp_Sh_Lpt_Sep was a temporary staging table used for a one-off task.
--      It has been COMPLETELY REMOVED from the view.
--   2. In HRIMS.HR_EMP_PIS, an employee may have multiple PC numbers (e.g. 
--      an old cadre PCNO and a current/promoted PCNO) under the same PIS.
--   3. When joining with V_GpfpPran_Max:
--      - Using TRIM(UPPER(...)) ensures blank-padding in CHAR columns or
--        whitespace discrepancies never cause join failures (which previously
--        caused PIS to return NULL for one of the PC numbers).
--      - The window function MAX(TRIM(g.ACCNO)) OVER (PARTITION BY TRIM(p.PIS))
--        guarantees that all PC numbers belonging to the same employee receive
--        the GPF/PRAN number, even if V_GpfpPran_Max only stores the GPF/PRAN
--        against the latest PC number.
-- ====================================================================

-- --------------------------------------------------------------------
-- OPTION 1: Using HRIMS.HR_EMP_PIS (Columns: PCNO, PIS) + V_GpfpPran_Max
-- Use this if your source table for PIS is HRIMS.HR_EMP_PIS
-- --------------------------------------------------------------------
CREATE OR REPLACE VIEW V_EMP_DETAILS AS
SELECT 
    COALESCE(TRIM(p.PCNO), TRIM(g.PCNO)) AS PCNO,
    TRIM(p.PIS) AS PIS,
    COALESCE(TRIM(g.ACCNO), MAX(TRIM(g.ACCNO)) OVER (PARTITION BY TRIM(p.PIS))) AS GPFPRAN
FROM HRIMS.HR_EMP_PIS p
FULL OUTER JOIN V_GpfpPran_Max g 
    ON TRIM(UPPER(p.PCNO)) = TRIM(UPPER(g.PCNO));

-- Permissions:
-- GRANT SELECT ON V_EMP_DETAILS TO <your_app_db_user>;


-- --------------------------------------------------------------------
-- OPTION 2: If Bank Account Number is also required from V_Emp_Pis_BankAccountDetails
-- Columns: PIS, PCNO, ACCNO (where ACCNO = Bank Account Number)
-- --------------------------------------------------------------------
-- CREATE OR REPLACE VIEW V_EMP_DETAILS AS
-- SELECT 
--     COALESCE(TRIM(p.PCNO), TRIM(g.PCNO)) AS PCNO,
--     TRIM(p.PIS) AS PIS,
--     COALESCE(TRIM(g.ACCNO), MAX(TRIM(g.ACCNO)) OVER (PARTITION BY TRIM(p.PIS))) AS GPFPRAN,
--     TRIM(p.ACCNO) AS BANK_ACCNO,
--     TRIM(p.ACCNO) AS ACCNO
-- FROM V_Emp_Pis_BankAccountDetails p
-- FULL OUTER JOIN V_GpfpPran_Max g 
--     ON TRIM(UPPER(p.PCNO)) = TRIM(UPPER(g.PCNO));
