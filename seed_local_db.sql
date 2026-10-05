-- ====================================================================
-- SEED SCRIPT FOR LOCAL ORACLE DATABASE (Oracle 21c XE)
-- Tables used (3 tables):
--   1. Temp_Sh_Lpt_Sep   (Columns: SLNO, GPFPRAN)
--   2. V_GpfpPran_Max     (Columns: ACCNO, PCNO)
--   3. HRIMS.HR_EMP_PIS   (Columns: PCNO, PIS)
-- Resulting New View:
--   V_EMP_DETAILS         (Columns: SLNO, GPFPRAN, ACCNO, PCNO, PIS)
-- ====================================================================

-- 1. Drop old view V_EMP_PisAccPcno if it exists (not used anymore)
BEGIN
    EXECUTE IMMEDIATE 'DROP VIEW V_EMP_PisAccPcno';
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

-- 4. Create Table 3: HRIMS.HR_EMP_PIS
BEGIN
    EXECUTE IMMEDIATE 'TRUNCATE TABLE HRIMS.HR_EMP_PIS';
EXCEPTION
    WHEN OTHERS THEN
        BEGIN
            EXECUTE IMMEDIATE 'DROP TABLE HRIMS.HR_EMP_PIS CASCADE CONSTRAINTS';
        EXCEPTION WHEN OTHERS THEN NULL;
        END;
        EXECUTE IMMEDIATE 'CREATE TABLE HRIMS.HR_EMP_PIS (PCNO VARCHAR2(50), PIS VARCHAR2(50))';
END;
/

BEGIN
    EXECUTE IMMEDIATE 'CREATE OR REPLACE SYNONYM HR_EMP_PIS FOR HRIMS.HR_EMP_PIS';
EXCEPTION
    WHEN OTHERS THEN NULL;
END;
/

-- ====================================================================
-- SEED DATA
-- ====================================================================

-- Seed Table 1: Temp_Sh_Lpt_Sep
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (1, 'GPF-1111');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (2, 'GPF-2222');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (3, 'PRAN-3333');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (4, 'PRAN-4444');
INSERT INTO Temp_Sh_Lpt_Sep (SLNO, GPFPRAN) VALUES (5, 'GPF-5555');

-- Seed Table 2: V_GpfpPran_Max
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-1111', '5001');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-2222', '5002');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-2222', '5010');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-3333', '5003');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('PRAN-4444', '5004');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-5555', '5015');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-5555', '5025');
INSERT INTO V_GpfpPran_Max (ACCNO, PCNO) VALUES ('GPF-9999', '5099');

-- Seed Table 3: HRIMS.HR_EMP_PIS
-- Case 1: PIS 2008AE10 has single PCNO 5001
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5001', '2008AE10');

-- Case 2: PIS 2008AE12 has TWO PCNOs (5002 and 5010) -> Selects 5010
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5002', '2008AE12');
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5010', '2008AE12');

-- Case 3: PIS 2012BC24 has single PCNO 5003
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5003', '2012BC24');

-- Case 4: PIS 2014CD36 has single PCNO 5004
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5004', '2014CD36');

-- Case 5: PIS 2015EF48 has TWO PCNOs (5015 and 5025) -> Selects 5025
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5015', '2015EF48');
INSERT INTO HRIMS.HR_EMP_PIS (PCNO, PIS) VALUES ('5025', '2015EF48');

COMMIT;

-- ====================================================================
-- CREATE NEW VIEW USING THE 3 TABLES
-- (FULL OUTER JOIN guarantees that lookups by GPFPRAN, ACCNO, PCNO, or PIS
--  all succeed even if an employee is present in only one or two tables)
-- ====================================================================
CREATE OR REPLACE VIEW V_EMP_DETAILS AS
SELECT 
    a.SLNO,
    COALESCE(a.GPFPRAN, b.ACCNO) AS GPFPRAN,
    COALESCE(b.ACCNO, a.GPFPRAN) AS ACCNO,
    COALESCE(b.PCNO, c.PCNO) AS PCNO,
    c.PIS
FROM Temp_Sh_Lpt_Sep a
FULL OUTER JOIN V_GpfpPran_Max b ON a.GPFPRAN = b.ACCNO
FULL OUTER JOIN HRIMS.HR_EMP_PIS c ON b.PCNO = c.PCNO;

-- Verify View
PROMPT === V_EMP_DETAILS View Created Successfully ===;
SELECT * FROM V_EMP_DETAILS ORDER BY PCNO;
EXIT;
