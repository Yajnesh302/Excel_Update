using System;
using System.IO;
using System.Data;
using System.Configuration;
using System.Web;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;

namespace ExcelProcessor
{
    public class EmpRecord
    {
        public string Pcno { get; set; }
        public string Pis { get; set; }
        public string GpfPran { get; set; }
        public string BankAccNo { get; set; }
        public string AccNo
        {
            get { return BankAccNo; }
            set { BankAccNo = value; }
        }
    }

    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string action = Request.QueryString["action"];
            if (string.IsNullOrEmpty(action))
            {
                // Serve standard ASPX page rendering (default behavior)
                return;
            }

            // Route AJAX actions, writing raw JSON or binary and stopping page cycle
            Response.Clear();
            Response.ContentType = "application/json";

            try
            {
                switch (action.ToLowerInvariant())
                {
                    case "status":
                        CheckStatus();
                        break;
                    case "sheets":
                    case "inspect":
                        InspectUploadedFile();
                        break;
                    case "process":
                        ProcessExcel();
                        break;
                    default:
                        ReturnError("Invalid action specified.");
                        break;
                }
            }
            catch (Exception ex)
            {
                ReturnError(ex.Message);
            }

            Response.Flush();
            Response.SuppressContent = true;
            HttpContext.Current.ApplicationInstance.CompleteRequest();
        }

        private string GetConfigTableName()
        {
            string tableName = ConfigurationManager.AppSettings["DatabaseTableName"];
            if (string.IsNullOrWhiteSpace(tableName))
            {
                tableName = "V_EMP_DETAILS";
            }
            return tableName.Trim().ToUpperInvariant();
        }

        private string GetConfigPcnoColumn()
        {
            string col = ConfigurationManager.AppSettings["DatabasePcnoColumn"];
            return string.IsNullOrWhiteSpace(col) ? "PCNO" : col.Trim().ToUpperInvariant();
        }

        private string GetConfigPisColumn()
        {
            string col = ConfigurationManager.AppSettings["DatabasePisColumn"];
            return string.IsNullOrWhiteSpace(col) ? "PIS" : col.Trim().ToUpperInvariant();
        }

        private string GetConfigGpfPranColumn()
        {
            string col = ConfigurationManager.AppSettings["DatabaseGpfPranColumn"];
            return string.IsNullOrWhiteSpace(col) ? "GPFPRAN" : col.Trim().ToUpperInvariant();
        }

        private string GetConfigBankAccNoColumn()
        {
            string col = ConfigurationManager.AppSettings["DatabaseBankAccNoColumn"];
            if (string.IsNullOrWhiteSpace(col))
            {
                col = ConfigurationManager.AppSettings["DatabaseAccNoColumn"];
            }
            return string.IsNullOrWhiteSpace(col) ? "BANK_ACCNO" : col.Trim().ToUpperInvariant();
        }

        private string GetConfigAccNoColumn()
        {
            string col = ConfigurationManager.AppSettings["DatabaseAccNoColumn"];
            return string.IsNullOrWhiteSpace(col) ? "ACCNO" : col.Trim().ToUpperInvariant();
        }

        private void CheckStatus()
        {
            string connStr = ConfigurationManager.ConnectionStrings["OracleConn"].ConnectionString;
            bool connected = false;
            bool tableExists = false;
            int recordCount = 0;
            string errorMsg = "";

            string tableName = GetConfigTableName();

            try
            {
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();
                    connected = true;

                    // Check both USER_VIEWS / USER_TABLES and ALL_VIEWS / ALL_TABLES
                    string checkSql = @"
                        SELECT COUNT(*) FROM (
                            SELECT VIEW_NAME AS OBJ_NAME FROM USER_VIEWS WHERE VIEW_NAME = :tName
                            UNION ALL
                            SELECT TABLE_NAME AS OBJ_NAME FROM USER_TABLES WHERE TABLE_NAME = :tName
                            UNION ALL
                            SELECT VIEW_NAME AS OBJ_NAME FROM ALL_VIEWS WHERE VIEW_NAME = :tName
                            UNION ALL
                            SELECT TABLE_NAME AS OBJ_NAME FROM ALL_TABLES WHERE TABLE_NAME = :tName
                        )";

                    using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                    {
                        cmd.Parameters.Add(new OracleParameter("tName", tableName));
                        int count = Convert.ToInt32(cmd.ExecuteScalar());
                        tableExists = count > 0;
                    }

                    if (tableExists)
                    {
                        string countSql = "SELECT COUNT(*) FROM " + tableName;
                        using (OracleCommand cmd = new OracleCommand(countSql, conn))
                        {
                            recordCount = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }

            var supportedKeys = new List<object>
            {
                new { id = "PCNO", name = "PC Number (PCNO)", description = "Employee Cadre / PC Number" },
                new { id = "PIS", name = "PIS Number (PIS)", description = "Personnel Information System Number" },
                new { id = "GPFPRAN", name = "GPF / PRAN (GPFPRAN)", description = "GPF / PRAN Account Number (from V_GpfpPran_Max)" },
                new { id = "BANK_ACCNO", name = "Bank Account No (ACCNO)", description = "Bank Account Number (from V_Emp_Pis_BankAccountDetails)" }
            };

            var result = new
            {
                connected = connected,
                tableExists = tableExists,
                tableName = tableName,
                recordCount = recordCount,
                supportedKeys = supportedKeys,
                error = errorMsg
            };

            var serializer = new JavaScriptSerializer();
            Response.Write(serializer.Serialize(result));
        }

        private void InspectUploadedFile()
        {
            if (Request.Files.Count == 0)
            {
                ReturnError("No file uploaded.");
                return;
            }

            HttpPostedFile file = Request.Files[0];
            if (file == null || file.ContentLength == 0)
            {
                ReturnError("Uploaded file is empty.");
                return;
            }

            try
            {
                List<string> sheetNames = new List<string>();
                List<object> columns = new List<object>();
                string detectedColumn = null;
                string detectedKeyType = null;

                if (file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    sheetNames.Add("Default");
                    List<List<string>> csvRows = ParseCsv(file.InputStream);
                    if (csvRows.Count > 0)
                    {
                        int headerRowIdx = FindHeaderRowCsv(csvRows);
                        var headerRow = csvRows[headerRowIdx];
                        for (int i = 0; i < headerRow.Count; i++)
                        {
                            string colName = headerRow[i];
                            if (string.IsNullOrWhiteSpace(colName)) continue;
                            string dType = DetectKeyType(colName);
                            columns.Add(new { index = i + 1, name = colName, detectedType = dType });
                            if (detectedKeyType == null && dType != null)
                            {
                                detectedColumn = colName;
                                detectedKeyType = dType;
                            }
                        }
                    }
                }
                else
                {
                    using (ExcelPackage package = new ExcelPackage(file.InputStream))
                    {
                        foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                        {
                            sheetNames.Add(ws.Name);
                        }

                        if (package.Workbook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets[1]; // 1-based index in EPPlus
                            var dim = ws.Dimension;
                            if (dim != null)
                            {
                                int rCount = dim.End.Row;
                                int cCount = dim.End.Column;
                                int maxScan = Math.Min(rCount, 25);
                                int headerRow = 1;

                                // Look for first row that contains at least one recognized key or non-empty cells
                                for (int r = 1; r <= maxScan; r++)
                                {
                                    bool hasAnyRecognized = false;
                                    for (int c = 1; c <= cCount; c++)
                                    {
                                        string text = ws.Cells[r, c].Text;
                                        if (DetectKeyType(text) != null)
                                        {
                                            hasAnyRecognized = true;
                                            break;
                                        }
                                    }
                                    if (hasAnyRecognized)
                                    {
                                        headerRow = r;
                                        break;
                                    }
                                }

                                for (int c = 1; c <= cCount; c++)
                                {
                                    string colName = ws.Cells[headerRow, c].Text;
                                    if (string.IsNullOrWhiteSpace(colName)) continue;
                                    string dType = DetectKeyType(colName);
                                    columns.Add(new { index = c, name = colName, detectedType = dType });
                                    if (detectedKeyType == null && dType != null)
                                    {
                                        detectedColumn = colName;
                                        detectedKeyType = dType;
                                    }
                                }
                            }
                        }
                    }
                }

                var serializer = new JavaScriptSerializer();
                Response.Write(serializer.Serialize(new
                {
                    success = true,
                    sheets = sheetNames,
                    columns = columns,
                    detectedColumn = detectedColumn,
                    detectedKeyType = detectedKeyType
                }));
            }
            catch (Exception ex)
            {
                ReturnError("Failed to inspect file: " + ex.Message);
            }
        }

        private void ProcessExcel()
        {
            if (Request.Files.Count == 0)
            {
                ReturnError("No file uploaded.");
                return;
            }

            HttpPostedFile file = Request.Files[0];
            if (file == null || file.ContentLength == 0)
            {
                ReturnError("Uploaded file is empty.");
                return;
            }

            // Input Column & Key Type
            string inputColName = Request.Form["inputCol"];
            string keyType = Request.Form["keyType"]; // "PCNO", "PIS", "GPFPRAN", "BANK_ACCNO", "ACCNO"

            if (string.IsNullOrWhiteSpace(keyType))
            {
                ReturnError("Identifier key type is required (PCNO, PIS, GPFPRAN, or BANK_ACCNO).");
                return;
            }
            keyType = keyType.Trim().ToUpperInvariant();
            if (keyType == "ACCNO") keyType = "BANK_ACCNO";

            // Output columns to append
            string outputColsParam = Request.Form["outputCols"];
            if (string.IsNullOrWhiteSpace(outputColsParam))
            {
                // Fallback: pick standard complementary output columns
                if (keyType == "PIS") outputColsParam = "PCNO,GPFPRAN,BANK_ACCNO";
                else if (keyType == "PCNO") outputColsParam = "PIS,GPFPRAN,BANK_ACCNO";
                else if (keyType == "GPFPRAN") outputColsParam = "PCNO,PIS,BANK_ACCNO";
                else outputColsParam = "PCNO,PIS,GPFPRAN";
            }

            string[] outputColsRaw = outputColsParam.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> outputCols = new List<string>();
            foreach (string oc in outputColsRaw)
            {
                string norm = oc.Trim().ToUpperInvariant();
                if ((norm == "PCNO" || norm == "PIS" || norm == "GPFPRAN" || norm == "BANK_ACCNO" || norm == "ACCNO") && !outputCols.Contains(norm))
                {
                    outputCols.Add(norm);
                }
            }

            if (outputCols.Count == 0)
            {
                ReturnError("Please select at least one output column to append.");
                return;
            }

            // Load data from Oracle View into memory
            Dictionary<string, EmpRecord> byPcno;
            Dictionary<string, EmpRecord> byPis;
            Dictionary<string, EmpRecord> byGpfPran;
            Dictionary<string, EmpRecord> byBankAccNo;

            string loadError = LoadViewData(out byPcno, out byPis, out byGpfPran, out byBankAccNo);
            if (!string.IsNullOrEmpty(loadError))
            {
                ReturnError("Records Error: " + loadError);
                return;
            }

            if (file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                ProcessCsvFile(file, inputColName, keyType, outputCols, byPcno, byPis, byGpfPran, byBankAccNo);
                return;
            }

            // Worksheets to process
            string selectedSheetsParam = Request.Form["sheets"];
            List<string> selectedSheetsList = new List<string>();
            if (!string.IsNullOrEmpty(selectedSheetsParam))
            {
                string[] ss = selectedSheetsParam.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string s in ss)
                {
                    string trimmed = s.Trim();
                    if (!string.IsNullOrEmpty(trimmed) && !selectedSheetsList.Contains(trimmed))
                    {
                        selectedSheetsList.Add(trimmed);
                    }
                }
            }

            try
            {
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        ReturnError("Excel file does not contain any worksheets.");
                        return;
                    }

                    bool processedAny = false;

                    foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                    {
                        if (selectedSheetsList.Count > 0 && !selectedSheetsList.Contains(ws.Name))
                        {
                            continue;
                        }

                        var dimension = ws.Dimension;
                        if (dimension == null)
                        {
                            if (selectedSheetsList.Contains(ws.Name))
                            {
                                ReturnError(string.Format("Selected sheet '{0}' is empty.", ws.Name));
                                return;
                            }
                            continue;
                        }

                        int rCount = dimension.End.Row;
                        int cCount = dimension.End.Column;
                        int maxHeaderScanRows = Math.Min(rCount, 25);
                        int inputColIndex = -1;
                        int headerRowIndex = -1;

                        // 1. Locate the header row and input column
                        for (int r = 1; r <= maxHeaderScanRows; r++)
                        {
                            for (int c = 1; c <= cCount; c++)
                            {
                                string cellText = ws.Cells[r, c].Text;
                                if (IsMatchingInputColumn(cellText, inputColName, keyType))
                                {
                                    inputColIndex = c;
                                    headerRowIndex = r;
                                    break;
                                }
                            }
                            if (inputColIndex != -1) break;
                        }

                        if (inputColIndex == -1)
                        {
                            // If user specified an explicit column name or index, try fallback
                            int parsedCol;
                            if (int.TryParse(inputColName, out parsedCol) && parsedCol >= 1 && parsedCol <= cCount)
                            {
                                inputColIndex = parsedCol;
                                headerRowIndex = 1;
                            }
                            else if (selectedSheetsList.Contains(ws.Name))
                            {
                                ReturnError(string.Format("Selected sheet '{0}' does not contain column '{1}' or key type '{2}'.", ws.Name, inputColName ?? keyType, keyType));
                                return;
                            }
                            else
                            {
                                continue;
                            }
                        }

                        processedAny = true;

                        // 2. Append chosen output column headers
                        Dictionary<string, int> targetColIndexes = new Dictionary<string, int>();
                        int currentColCount = cCount;
                        foreach (string outCol in outputCols)
                        {
                            currentColCount++;
                            targetColIndexes[outCol] = currentColCount;
                            ws.Cells[headerRowIndex, currentColCount].Value = GetOutputColumnHeaderTitle(outCol);
                        }

                        // 3. Process each row
                        for (int r = headerRowIndex + 1; r <= rCount; r++)
                        {
                            string rawVal = ws.Cells[r, inputColIndex].Text;
                            string cleanedVal = CleanValue(rawVal);

                            if (string.IsNullOrEmpty(cleanedVal))
                            {
                                // Leave empty
                                continue;
                            }

                            EmpRecord matchedRecord = LookupRecord(cleanedVal, keyType, byPcno, byPis, byGpfPran, byBankAccNo);

                            foreach (string outCol in outputCols)
                            {
                                int targetCol = targetColIndexes[outCol];
                                string valToInsert = GetRecordValue(matchedRecord, outCol);
                                ws.Cells[r, targetCol].Value = valToInsert;
                            }
                        }
                    }

                    if (!processedAny)
                    {
                        ReturnError("No valid worksheets were processed.");
                        return;
                    }

                    byte[] fileBytes;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        package.SaveAs(ms);
                        fileBytes = ms.ToArray();
                    }

                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename=\"{0}\"", file.FileName));
                    Response.BinaryWrite(fileBytes);
                    Response.Flush();
                    Response.SuppressContent = true;
                    HttpContext.Current.ApplicationInstance.CompleteRequest();
                }
            }
            catch (Exception ex)
            {
                ReturnError("Excel Processing Error: " + ex.Message);
            }
        }

        private void ProcessCsvFile(HttpPostedFile file, string inputColName, string keyType, List<string> outputCols,
            Dictionary<string, EmpRecord> byPcno, Dictionary<string, EmpRecord> byPis, Dictionary<string, EmpRecord> byGpfPran, Dictionary<string, EmpRecord> byBankAccNo)
        {
            try
            {
                List<List<string>> csvRows = ParseCsv(file.InputStream);
                if (csvRows.Count == 0)
                {
                    ReturnError("Uploaded CSV file is empty.");
                    return;
                }

                int inputColIndex = -1;
                int headerRowIndex = -1;
                int rowCount = csvRows.Count;
                int colCount = 0;

                int maxHeaderScanRows = Math.Min(rowCount, 25);
                for (int r = 0; r < maxHeaderScanRows; r++)
                {
                    var row = csvRows[r];
                    for (int c = 0; c < row.Count; c++)
                    {
                        string cellText = row[c];
                        if (IsMatchingInputColumn(cellText, inputColName, keyType))
                        {
                            inputColIndex = c;
                            headerRowIndex = r;
                            colCount = row.Count;
                            break;
                        }
                    }
                    if (inputColIndex != -1) break;
                }

                if (inputColIndex == -1)
                {
                    ReturnError(string.Format("Could not find identifier column '{0}' in the uploaded CSV file.", inputColName ?? keyType));
                    return;
                }

                var headerRow = csvRows[headerRowIndex];
                foreach (string outCol in outputCols)
                {
                    headerRow.Add(GetOutputColumnHeaderTitle(outCol));
                }

                for (int r = headerRowIndex + 1; r < rowCount; r++)
                {
                    var row = csvRows[r];
                    while (row.Count < colCount)
                    {
                        row.Add("");
                    }

                    string rawVal = (inputColIndex < row.Count) ? row[inputColIndex] : "";
                    string cleanedVal = CleanValue(rawVal);

                    EmpRecord matchedRecord = string.IsNullOrEmpty(cleanedVal)
                        ? null
                        : LookupRecord(cleanedVal, keyType, byPcno, byPis, byGpfPran, byBankAccNo);

                    foreach (string outCol in outputCols)
                    {
                        string valToInsert = string.IsNullOrEmpty(cleanedVal) ? "" : GetRecordValue(matchedRecord, outCol);
                        row.Add(valToInsert);
                    }
                }

                byte[] fileBytes = WriteCsv(csvRows);

                Response.Clear();
                Response.ContentType = "text/csv";
                Response.AddHeader("Content-Disposition", string.Format("attachment; filename=\"{0}\"", file.FileName));
                Response.BinaryWrite(fileBytes);
                Response.Flush();
                Response.SuppressContent = true;
                HttpContext.Current.ApplicationInstance.CompleteRequest();
            }
            catch (Exception ex)
            {
                ReturnError("CSV Processing Error: " + ex.Message);
            }
        }

        private static bool HasColumn(OracleDataReader reader, string columnName)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (string.Equals(reader.GetName(i), columnName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private string LoadViewData(out Dictionary<string, EmpRecord> byPcno, 
                                   out Dictionary<string, EmpRecord> byPis, 
                                   out Dictionary<string, EmpRecord> byGpfPran, 
                                   out Dictionary<string, EmpRecord> byBankAccNo)
        {
            byPcno = new Dictionary<string, EmpRecord>(StringComparer.OrdinalIgnoreCase);
            byPis = new Dictionary<string, EmpRecord>(StringComparer.OrdinalIgnoreCase);
            byGpfPran = new Dictionary<string, EmpRecord>(StringComparer.OrdinalIgnoreCase);
            byBankAccNo = new Dictionary<string, EmpRecord>(StringComparer.OrdinalIgnoreCase);

            string tableName = GetConfigTableName();
            string pcnoCol = GetConfigPcnoColumn();
            string pisCol = GetConfigPisColumn();
            string gpfpranCol = GetConfigGpfPranColumn();
            string bankAccCol = GetConfigBankAccNoColumn();
            string accnoCol = GetConfigAccNoColumn();

            string connStr = ConfigurationManager.ConnectionStrings["OracleConn"].ConnectionString;

            try
            {
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    // Detect actual columns in the view/table
                    HashSet<string> existingCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        using (OracleCommand schemaCmd = new OracleCommand("SELECT * FROM " + tableName + " WHERE 1 = 0", conn))
                        using (OracleDataReader schemaReader = schemaCmd.ExecuteReader(CommandBehavior.SchemaOnly))
                        {
                            var schemaTable = schemaReader.GetSchemaTable();
                            if (schemaTable != null)
                            {
                                foreach (DataRow row in schemaTable.Rows)
                                {
                                    string cName = row["ColumnName"] != null ? row["ColumnName"].ToString() : "";
                                    if (!string.IsNullOrEmpty(cName))
                                    {
                                        existingCols.Add(cName.ToUpperInvariant());
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Fallback if schema only query fails
                    }

                    List<string> selectCols = new List<string>();
                    string actualPcnoCol = existingCols.Count == 0 || existingCols.Contains(pcnoCol) ? pcnoCol : null;
                    string actualPisCol = existingCols.Count == 0 || existingCols.Contains(pisCol) ? pisCol : null;
                    string actualGpfPranCol = existingCols.Count == 0 || existingCols.Contains(gpfpranCol) ? gpfpranCol : null;

                    string actualBankAccCol = null;
                    if (existingCols.Count == 0) actualBankAccCol = bankAccCol;
                    else if (existingCols.Contains(bankAccCol)) actualBankAccCol = bankAccCol;
                    else if (existingCols.Contains(accnoCol)) actualBankAccCol = accnoCol;

                    if (actualPcnoCol != null && !selectCols.Contains(actualPcnoCol)) selectCols.Add(actualPcnoCol);
                    if (actualPisCol != null && !selectCols.Contains(actualPisCol)) selectCols.Add(actualPisCol);
                    if (actualGpfPranCol != null && !selectCols.Contains(actualGpfPranCol)) selectCols.Add(actualGpfPranCol);
                    if (actualBankAccCol != null && !selectCols.Contains(actualBankAccCol)) selectCols.Add(actualBankAccCol);

                    if (selectCols.Count == 0) selectCols.Add("*");

                    string query = "SELECT " + string.Join(", ", selectCols) + " FROM " + tableName;

                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string pcno = "";
                            string pis = "";
                            string gpfpran = "";
                            string bankAccNo = "";

                            if (actualPcnoCol != null && HasColumn(reader, actualPcnoCol))
                                pcno = CleanValue(reader[actualPcnoCol] != DBNull.Value ? reader[actualPcnoCol].ToString() : "");

                            if (actualPisCol != null && HasColumn(reader, actualPisCol))
                                pis = CleanValue(reader[actualPisCol] != DBNull.Value ? reader[actualPisCol].ToString() : "");

                            if (actualGpfPranCol != null && HasColumn(reader, actualGpfPranCol))
                                gpfpran = CleanValue(reader[actualGpfPranCol] != DBNull.Value ? reader[actualGpfPranCol].ToString() : "");

                            if (actualBankAccCol != null && HasColumn(reader, actualBankAccCol))
                                bankAccNo = CleanValue(reader[actualBankAccCol] != DBNull.Value ? reader[actualBankAccCol].ToString() : "");

                            var record = new EmpRecord
                            {
                                Pcno = pcno,
                                Pis = pis,
                                GpfPran = gpfpran,
                                BankAccNo = bankAccNo
                            };

                            // Map by PCNO
                            if (!string.IsNullOrEmpty(pcno))
                            {
                                EmpRecord existing;
                                if (byPcno.TryGetValue(pcno, out existing))
                                {
                                    if (string.IsNullOrEmpty(existing.Pis) && !string.IsNullOrEmpty(record.Pis))
                                        existing.Pis = record.Pis;
                                    if (string.IsNullOrEmpty(existing.GpfPran) && !string.IsNullOrEmpty(record.GpfPran))
                                        existing.GpfPran = record.GpfPran;
                                    if (string.IsNullOrEmpty(existing.BankAccNo) && !string.IsNullOrEmpty(record.BankAccNo))
                                        existing.BankAccNo = record.BankAccNo;
                                }
                                else
                                {
                                    byPcno[pcno] = record;
                                }
                            }

                            // Map by PIS (Apply MAX(PCNO) rule: if a PIS has multiple PCNOs, pick the MAX PCNO)
                            if (!string.IsNullOrEmpty(pis))
                            {
                                EmpRecord existing;
                                if (byPis.TryGetValue(pis, out existing))
                                {
                                    if (ComparePcno(record.Pcno, existing.Pcno) > 0)
                                    {
                                        byPis[pis] = record;
                                    }
                                }
                                else
                                {
                                    byPis[pis] = record;
                                }
                            }

                            // Map by GPFPRAN (Apply MAX(PCNO) rule)
                            if (!string.IsNullOrEmpty(gpfpran))
                            {
                                EmpRecord existing;
                                if (byGpfPran.TryGetValue(gpfpran, out existing))
                                {
                                    if (ComparePcno(record.Pcno, existing.Pcno) > 0)
                                    {
                                        byGpfPran[gpfpran] = record;
                                    }
                                }
                                else
                                {
                                    byGpfPran[gpfpran] = record;
                                }
                            }

                            // Map by Bank Account Number / ACCNO (Apply MAX(PCNO) rule)
                            if (!string.IsNullOrEmpty(bankAccNo))
                            {
                                EmpRecord existing;
                                if (byBankAccNo.TryGetValue(bankAccNo, out existing))
                                {
                                    if (ComparePcno(record.Pcno, existing.Pcno) > 0)
                                    {
                                        byBankAccNo[bankAccNo] = record;
                                    }
                                }
                                else
                                {
                                    byBankAccNo[bankAccNo] = record;
                                }
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private static EmpRecord LookupRecord(string cleanedVal, string keyType,
            Dictionary<string, EmpRecord> byPcno, 
            Dictionary<string, EmpRecord> byPis, 
            Dictionary<string, EmpRecord> byGpfPran, 
            Dictionary<string, EmpRecord> byBankAccNo)
        {
            if (string.IsNullOrEmpty(cleanedVal)) return null;

            EmpRecord record = null;
            switch (keyType)
            {
                case "PCNO":
                    byPcno.TryGetValue(cleanedVal, out record);
                    break;
                case "PIS":
                    byPis.TryGetValue(cleanedVal, out record);
                    break;
                case "GPFPRAN":
                case "GPF":
                case "PRAN":
                    byGpfPran.TryGetValue(cleanedVal, out record);
                    break;
                case "BANK_ACCNO":
                case "ACCNO":
                case "BANKACCNO":
                case "BANK_ACCOUNT":
                    byBankAccNo.TryGetValue(cleanedVal, out record);
                    break;
            }
            return record;
        }

        private static string GetRecordValue(EmpRecord record, string targetCol)
        {
            if (record == null) return "Not Found";

            string val = "";
            switch (targetCol)
            {
                case "PCNO":
                    val = record.Pcno;
                    break;
                case "PIS":
                    val = record.Pis;
                    break;
                case "GPFPRAN":
                case "GPF":
                case "PRAN":
                    val = record.GpfPran;
                    break;
                case "BANK_ACCNO":
                case "ACCNO":
                case "BANKACCNO":
                case "BANK_ACCOUNT":
                    val = record.BankAccNo;
                    break;
            }

            return string.IsNullOrEmpty(val) ? "Not Found" : val;
        }

        private static string GetOutputColumnHeaderTitle(string colKey)
        {
            switch (colKey)
            {
                case "PCNO": return "PCNO";
                case "PIS": return "PIS";
                case "GPFPRAN": return "GPFPRAN";
                case "BANK_ACCNO": return "BANK_ACCNO";
                case "ACCNO": return "ACCNO";
                default: return colKey;
            }
        }

        private static int ComparePcno(string a, string b)
        {
            if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return 0;
            if (string.IsNullOrEmpty(a)) return -1;
            if (string.IsNullOrEmpty(b)) return 1;

            long numA, numB;
            if (long.TryParse(a.Trim(), out numA) && long.TryParse(b.Trim(), out numB))
            {
                return numA.CompareTo(numB);
            }

            var matchA = Regex.Match(a, @"\d+");
            var matchB = Regex.Match(b, @"\d+");
            if (matchA.Success && matchB.Success && long.TryParse(matchA.Value, out numA) && long.TryParse(matchB.Value, out numB))
            {
                int cmp = numA.CompareTo(numB);
                if (cmp != 0) return cmp;
            }

            return string.Compare(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static string DetectKeyType(string headerName)
        {
            if (string.IsNullOrWhiteSpace(headerName)) return null;

            string norm = NormalizeHeader(headerName);

            // Check GPF / PRAN first (in V_GpfpPran_Max)
            if (norm == "GPFPRAN" || norm == "GPF_PRAN" || norm == "GPF/PRAN" || norm == "GPF PRAN" ||
                norm == "GPF" || norm == "PRAN" || norm == "GPFNO" || norm == "PRANNO" ||
                norm == "GPF NO" || norm == "PRAN NO" || norm == "GPF_NO" || norm == "PRAN_NO" ||
                norm == "GPF/PRAN NO" || norm == "GPF / PRAN NO")
            {
                return "GPFPRAN";
            }

            // Check Bank Account (in V_Emp_Pis_BankAccountDetails ACCNO column is Bank Account Number)
            if (norm == "BANK_ACCNO" || norm == "BANK ACC NO" || norm == "BANK_ACC_NO" || norm == "BANK ACCNO" ||
                norm == "BANK ACCOUNT" || norm == "BANK ACCOUNT NO" || norm == "BANK_ACCOUNT_NO" || norm == "BANK ACCOUNT NUMBER" ||
                norm == "BANK A/C" || norm == "BANK A/C NO" || norm == "BANK A/C NUMBER" || norm == "BANK_AC_NO" ||
                norm == "ACCNO" || norm == "ACC NO" || norm == "ACC_NO" || norm == "ACC NO." ||
                norm == "ACCOUNT" || norm == "ACCOUNTNO" || norm == "ACCOUNT NO" || norm == "ACCOUNT_NO" ||
                norm == "ACCOUNTNUMBER" || norm == "ACCOUNT_NUMBER" || norm == "ACCOUNT NUMBER" ||
                norm == "ACNO" || norm == "AC NO" || norm == "AC_NO" || norm == "A/C NO" || norm == "A/C NUMBER")
            {
                return "BANK_ACCNO";
            }

            // Check PIS
            if (norm == "PIS" || norm == "PISNO" || norm == "PIS NO" || norm == "PIS NO." || 
                norm == "PIS_NO" || norm == "PISNUMBER" || norm == "PIS NUMBER" || norm == "PIS_NUMBER" || 
                norm == "PISNUM" || norm == "PIS_NUM")
            {
                return "PIS";
            }

            // Check PCNO
            if (norm == "PCNO" || norm == "PC NO" || norm == "PC NO." || norm == "PC_NO" || 
                norm == "PCNUMBER" || norm == "PC NUMBER" || norm == "PC_NUMBER" || 
                norm == "PCNUM" || norm == "PC_NUM" ||
                norm == "EMPCODE" || norm == "EMP CODE" || norm == "EMP_CODE" || norm == "EMPLOYEE CODE")
            {
                return "PCNO";
            }

            return null;
        }

        private static bool IsMatchingInputColumn(string cellText, string specifiedColName, string keyType)
        {
            if (string.IsNullOrWhiteSpace(cellText)) return false;

            string normCell = NormalizeHeader(cellText);

            // If user explicitly specified an input column name from dropdown
            if (!string.IsNullOrWhiteSpace(specifiedColName))
            {
                string normSpecified = NormalizeHeader(specifiedColName);
                if (normCell == normSpecified) return true;
            }

            // Otherwise, check if it matches the keyType
            string detected = DetectKeyType(cellText);
            if (detected != null)
            {
                if (string.Equals(detected, keyType, StringComparison.OrdinalIgnoreCase)) return true;
                if ((keyType == "BANK_ACCNO" || keyType == "ACCNO") && (detected == "BANK_ACCNO" || detected == "ACCNO")) return true;
                if ((keyType == "GPFPRAN" || keyType == "GPF" || keyType == "PRAN") && (detected == "GPFPRAN")) return true;
            }

            return false;
        }

        private static string NormalizeHeader(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string cleaned = text.Replace('\u00A0', ' ')
                                 .Replace('\r', ' ')
                                 .Replace('\n', ' ')
                                 .Replace('\t', ' ');
            while (cleaned.Contains("  "))
            {
                cleaned = cleaned.Replace("  ", " ");
            }
            return cleaned.Trim().ToUpperInvariant();
        }

        private static string CleanValue(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string cleaned = text.Replace('\u00A0', ' ')
                                 .Replace('\r', ' ')
                                 .Replace('\n', ' ')
                                 .Replace('\t', ' ');
            while (cleaned.Contains("  "))
            {
                cleaned = cleaned.Replace("  ", " ");
            }
            return cleaned.Trim();
        }

        private static int FindHeaderRowCsv(List<List<string>> rows)
        {
            int maxScan = Math.Min(rows.Count, 25);
            for (int r = 0; r < maxScan; r++)
            {
                foreach (string cell in rows[r])
                {
                    if (DetectKeyType(cell) != null) return r;
                }
            }
            return 0;
        }

        private static List<List<string>> ParseCsv(Stream stream)
        {
            var result = new List<List<string>>();
            using (var reader = new StreamReader(stream, System.Text.Encoding.UTF8))
            {
                var currentRow = new List<string>();
                var currentField = new System.Text.StringBuilder();
                bool inQuotes = false;
                int ch;

                while ((ch = reader.Read()) != -1)
                {
                    char c = (char)ch;

                    if (inQuotes)
                    {
                        if (c == '"')
                        {
                            int nextCh = reader.Peek();
                            if (nextCh == '"')
                            {
                                currentField.Append('"');
                                reader.Read();
                            }
                            else
                            {
                                inQuotes = false;
                            }
                        }
                        else
                        {
                            currentField.Append(c);
                        }
                    }
                    else
                    {
                        if (c == '"')
                        {
                            inQuotes = true;
                        }
                        else if (c == ',')
                        {
                            currentRow.Add(currentField.ToString());
                            currentField.Length = 0;
                        }
                        else if (c == '\r')
                        {
                            if (reader.Peek() == '\n')
                            {
                                reader.Read();
                            }
                            currentRow.Add(currentField.ToString());
                            currentField.Length = 0;
                            result.Add(currentRow);
                            currentRow = new List<string>();
                        }
                        else if (c == '\n')
                        {
                            currentRow.Add(currentField.ToString());
                            currentField.Length = 0;
                            result.Add(currentRow);
                            currentRow = new List<string>();
                        }
                        else
                        {
                            currentField.Append(c);
                        }
                    }
                }

                if (currentField.Length > 0 || currentRow.Count > 0)
                {
                    currentRow.Add(currentField.ToString());
                    result.Add(currentRow);
                }
            }

            return result;
        }

        private static byte[] WriteCsv(List<List<string>> rows)
        {
            using (var ms = new MemoryStream())
            using (var writer = new StreamWriter(ms, System.Text.Encoding.UTF8))
            {
                foreach (var row in rows)
                {
                    for (int i = 0; i < row.Count; i++)
                    {
                        string field = row[i] ?? "";
                        if (field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n"))
                        {
                            writer.Write("\"" + field.Replace("\"", "\"\"") + "\"");
                        }
                        else
                        {
                            writer.Write(field);
                        }

                        if (i < row.Count - 1)
                        {
                            writer.Write(",");
                        }
                    }
                    writer.Write("\r\n");
                }
                writer.Flush();
                return ms.ToArray();
            }
        }

        private void ReturnError(string message)
        {
            Response.Clear();
            Response.ContentType = "application/json";
            Response.StatusCode = 200;
            var serializer = new JavaScriptSerializer();
            Response.Write(serializer.Serialize(new { error = message }));
        }
    }
}
