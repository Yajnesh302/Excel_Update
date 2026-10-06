-- ====================================================================
-- SEED SCRIPT FOR LOCAL ORACLE DATABASE (Oracle 21c XE)
-- Tables used (3 tables):
--   1. Temp_Sh_Lpt_Sep              (Columns: SLNO, GPFPRAN)
--   2. V_GpfpPran_Max                (Columns: ACCNO, PCNO)
--      [ACCNO refers to GPFPRAN]
--   3. V_Emp_Pis_BankAccountDetails  (Columns: PIS, PCNO, ACCNO)
--      [ACCNO refers to Bank Account Number, NOT GPFPRAN]
--
-- Resulting New View:
--   V_EMP_DETAILS (Columns: SLNO, GPFPRAN, PCNO, PIS, BANK_ACCNO, ACCNO)
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

-- 2. Create Table 1: Temp_Sh_Lpt_Sep
BEGIN
    EXECUTE IMMEDIATE 'DROP TABLE Temp_Sh_Lpt_Sep CASCADE CONSTRAINTS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

CREATE TABLE Temp_Sh_Lpt_Sep (
    SLNO NUMBER,
    GPFPRAN VARCHAR2(50)
);

-- 3. Create Table 2: V_GpfpPran_Max
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

-- 4. Create Table 3: V_Emp_Pis_BankAccountDetails
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

-- Seed Table 1: Temp_Sh_Lpt_Sep (GPFPRAN)
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (1, 'GPF-1111');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (2, 'GPF-2222');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (3, 'PRAN-3333');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (4, 'PRAN-4444');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (5, 'GPF-5555');

-- Seed Table 2: V_GpfpPran_Max (ACCNO = GPFPRAN, PCNO)
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-1111', '5001');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-2222', '5002');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-2222', '5010');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-3333', '5003');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-4444', '5004');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-5555', '5015');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-5555', '5025');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-9999', '5099');

-- Seed Table 3: V_Emp_Pis_BankAccountDetails (PIS, PCNO, ACCNO = Bank Account Number)
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
-- CREATE UNIFIED VIEW USING THE 3 TABLES
-- ====================================================================
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

-- Verify View
PROMPT === V_EMP_DETAILS View Created Successfully ===;
SELECT * FROM V_EMP_DETAILS ORDER BY PCNO;
EXIT;
