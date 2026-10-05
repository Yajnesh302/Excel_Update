-- ====================================================================
-- SCRIPT FOR COMPANY DATABASE ADMINISTRATOR
-- Purpose: Create or replace the unified View V_EMP_DETAILS
-- Using the 3 source tables:
--   1. Temp_Sh_Lpt_Sep (Columns: SLNO, GpfPran)
--   2. V_GpfpPran_Max   (Columns: ACCNO, PCNO)
--   3. Hrims.HR_Emp_Pis (Columns: Pcno, PIS)
--
-- Resulting View:
--   V_EMP_DETAILS (Columns: SLNO, GPFPRAN, ACCNO, PCNO, PIS)
-- (Note: V_EMP_PisAccPcno is NOT used)
-- ====================================================================

-- OPTION A (RECOMMENDED): FULL OUTER JOIN
-- Ensures lookups by any key (GPFPRAN, ACCNO, PCNO, PIS) succeed,
-- linking all 3 tables seamlessly.

CREATE OR REPLACE VIEW V_EMP_DETAILS AS
SELECT 
    a.SLNO,
    COALESCE(a.GPFPRAN, b.ACCNO) AS GPFPRAN,
    COALESCE(b.ACCNO, a.GPFPRAN) AS ACCNO,
    COALESCE(b.PCNO, c.PCNO) AS PCNO,
    c.PIS
FROM Temp_Sh_Lpt_Sep a
FULL OUTER JOIN V_GpfpPran_Max b ON a.GPFPRAN = b.ACCNO
FULL OUTER JOIN Hrims.HR_Emp_Pis c ON b.PCNO = c.PCNO;

-- --------------------------------------------------------------------
-- OPTION B: Standard LEFT JOIN (Matching the existing query pattern)
-- If your workflow always starts from Temp_Sh_Lpt_Sep:
-- --------------------------------------------------------------------
-- CREATE OR REPLACE VIEW V_EMP_DETAILS AS
-- SELECT 
--     a.SLNO,
--     a.GPFPRAN,
--     b.ACCNO,
--     b.PCNO,
--     c.PIS
-- FROM Temp_Sh_Lpt_Sep a
-- LEFT JOIN V_GpfpPran_Max b ON a.GPFPRAN = b.ACCNO
-- LEFT JOIN Hrims.HR_Emp_Pis c ON b.PCNO = c.PCNO;

-- Permissions:
-- GRANT SELECT ON V_EMP_DETAILS TO <your_app_db_user>;
