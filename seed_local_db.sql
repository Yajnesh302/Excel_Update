-- ====================================================================
-- SEED SCRIPT FOR LOCAL ORACLE DATABASE (Oracle 21c XE)
-- Permanent Master Tables/Views used:
--   1. V_GpfpPran_Max                (Columns: ACCNO, PCNO)
--      [ACCNO refers to GPFPRAN]
--   2. V_Emp_Pis_BankAccountDetails  (Columns: PIS, PCNO, ACCNO)
--      [ACCNO refers to Bank Account Number, NOT GPFPRAN]
--
-- Note: Temp_Sh_Lpt_Sep was a temporary staging table and is NOT used.
--
-- Resulting Unified View:
--   V_EMP_DETAILS (Columns: PCNO, PIS, GPFPRAN, BANK_ACCNO, ACCNO)
-- ====================================================================

-- 1. Drop old tables/views if they exist
BEGIN
    EXECUTE IMMEDIATE 'DROP VIEW V_EMP_PisAccPcno';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE HRIMS.HR_EMP_PIS CASCADE CONSTRAINTS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

-- Drop temporary table Temp_Sh_Lpt_Sep (not used)
BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE Temp_Sh_Lpt_Sep CASCADE CONSTRAINTS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

-- 2. Create Table 1: V_GpfpPran_Max (ACCNO = GPFPRAN, PCNO)
BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE V_GpfpPran_Max CASCADE CONSTRAINTS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

CREATE TABLE V_GpfpPran_Max (
    ACCNO VARCHAR2(50),
    PCNO VARCHAR2(50)
);

-- 3. Create Table 2: V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO = Bank Account)
BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE V_Emp_Pis_BankAccountDetails CASCADE CONSTRAINTS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

CREATE TABLE V_Emp_Pis_BankAccountDetails (
    PIS VARCHAR2(50),
    PCNO VARCHAR2(50),
    ACCNO VARCHAR2(50)
);

-- ====================================================================
-- SEED DATA
-- ====================================================================

-- Seed Table 1: V_GpfpPran_Max (ACCNO = GPFPRAN, PCNO)
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-1111', '5001');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-2222', '5010');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-3333', '5003');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-4444', '5004');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-5555', '5025');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-9999', '5099');

-- Seed Table 2: V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO = Bank Account Number)
-- Case 1: PIS 2008AE10 -> PCNO 5001, Bank A/C 10000000001
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2008AE10', '5001', '10000000001');

-- Case 2: PIS 2008AE12 -> TWO PCNOs (5002 & 5010), Max PCNO is 5010 with Bank A/C 10000000010
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2008AE12', '5002', '10000000002');
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2008AE12', '5010', '10000000010');

-- Case 3: PIS 2012BC24 -> PCNO 5003, Bank A/C 10000000003
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2012BC24', '5003', '10000000003');

-- Case 4: PIS 2014CD36 -> PCNO 5004, Bank A/C 10000000004
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2014CD36', '5004', '10000000004');

-- Case 5: PIS 2015EF48 -> TWO PCNOs (5015 & 5025), Max PCNO is 5025 with Bank A/C 10000000025
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2015EF48', '5015', '10000000015');
INSERT INTO V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO) VALUES ('2015EF48', '5025', '10000000025');

COMMIT;

-- ====================================================================
-- CREATE UNIFIED VIEW
-- ====================================================================
CREATE OR REPLACE VIEW V_EMP_DETAILS AS
SELECT 
    COALESCE(TRIM(g.PCNO), TRIM(p.PCNO)) AS PCNO,
    TRIM(p.PIS) AS PIS,
    COALESCE(TRIM(g.ACCNO), MAX(TRIM(g.ACCNO)) OVER (PARTITION BY TRIM(p.PIS))) AS GPFPRAN,
    TRIM(p.ACCNO) AS BANK_ACCNO,
    TRIM(p.ACCNO) AS ACCNO
FROM V_Emp_Pis_BankAccountDetails p
FULL OUTER JOIN V_GpfpPran_Max g 
    ON TRIM(UPPER(p.PCNO)) = TRIM(UPPER(g.PCNO));

-- ====================================================================
-- CREATE EXCEL_APPUSERS TABLE (Authorized Access)
-- ====================================================================
BEGIN
    EXECUTE IMMEDIATE '
    CREATE TABLE Excel_AppUsers (
        PCNO VARCHAR2(50) NOT NULL,
        Name VARCHAR2(200),
        Role NUMBER(1) DEFAULT 1 NOT NULL,
        DivName VARCHAR2(100),
        CONSTRAINT PK_Excel_AppUsers PRIMARY KEY (PCNO)
    )';
EXCEPTION
    WHEN OTHERS THEN
        BEGIN
            EXECUTE IMMEDIATE 'ALTER TABLE Excel_AppUsers ADD DivName VARCHAR2(100)';
        EXCEPTION
            WHEN OTHERS THEN NULL;
        END;
END;
/

-- Ensure default admin and test users exist
MERGE INTO Excel_AppUsers u
USING (
    SELECT '1001' AS PCNO, 'Admin User' AS Name, 4 AS Role, 'DKRM/ITISG' AS DivName FROM dual UNION ALL
    SELECT '1002' AS PCNO, 'Test User' AS Name, 1 AS Role, 'D-ADMIN/STORE' AS DivName FROM dual
) src
ON (u.PCNO = src.PCNO)
WHEN MATCHED THEN
    UPDATE SET u.DivName = src.DivName
WHEN NOT MATCHED THEN
    INSERT (PCNO, Name, Role, DivName) VALUES (src.PCNO, src.Name, src.Role, src.DivName);

COMMIT;

-- Verify View and Users
PROMPT === V_EMP_DETAILS View Created Successfully ===;
SELECT * FROM V_EMP_DETAILS ORDER BY PCNO;

PROMPT === Excel_AppUsers Authorized Accounts ===;
SELECT * FROM Excel_AppUsers WHERE PCNO IN ('1001', '1002');
EXIT;
