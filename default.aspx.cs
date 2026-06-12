using System;
using System.IO;
using System.Data;
using System.Configuration;
using System.Web;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;

namespace ExcelProcessor
{
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
                switch (action)
                {
                    case "status":
                        CheckStatus();
                        break;
                    case "sheets":
                        GetSheets();
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

        private static Dictionary<string, string> GetMappableColumnsConfig()
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string configStr = ConfigurationManager.AppSettings["DatabaseMappableColumns"];
            if (!string.IsNullOrEmpty(configStr))
            {
                string[] parts = configStr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string part in parts)
                {
                    string[] kv = part.Split(new char[] { ':' }, 2);
                    if (kv.Length == 2)
                    {
                        string dbCol = kv[0].Trim().ToUpper();
                        string displayName = kv[1].Trim();
                        if (!string.IsNullOrEmpty(dbCol) && !string.IsNullOrEmpty(displayName))
                        {
                            dict[dbCol] = displayName;
                        }
                    }
                }
            }
            if (dict.Count == 0)
            {
                dict["ACCOUNT_NUMBER"] = "Account Number";
                dict["NAME"] = "Name";
                dict["RANK"] = "Rank";
            }
            return dict;
        }

        private void CheckStatus()
        {
            string connStr = ConfigurationManager.ConnectionStrings["OracleConn"].ConnectionString;
            bool connected = false;
            bool tableExists = false;
            int recordCount = 0;
            string errorMsg = "";

            string tableName = ConfigurationManager.AppSettings["DatabaseTableName"];
            if (string.IsNullOrEmpty(tableName))
            {
                tableName = "EMPLOYEE";
            }
            tableName = tableName.Trim().ToUpper();

            try
            {
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();
                    connected = true;

                    // Query USER_TABLES in Oracle to verify existence of table
                    string checkTableSql = "SELECT COUNT(*) FROM USER_TABLES WHERE TABLE_NAME = :tableName";
                    using (OracleCommand cmd = new OracleCommand(checkTableSql, conn))
                    {
                        cmd.Parameters.Add(new OracleParameter("tableName", tableName));
                        int tableCount = Convert.ToInt32(cmd.ExecuteScalar());
                        tableExists = tableCount > 0;
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

            var mappableColsList = new List<object>();
            var configCols = GetMappableColumnsConfig();
            foreach (var kvp in configCols)
            {
                mappableColsList.Add(new { dbColumn = kvp.Key, displayName = kvp.Value });
            }

            var result = new
            {
                connected = connected,
                tableExists = tableExists,
                recordCount = recordCount,
                mappableColumns = mappableColsList,
                error = errorMsg
            };

            var serializer = new JavaScriptSerializer();
            Response.Write(serializer.Serialize(result));
        }

        private void GetSheets()
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
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    List<string> sheetNames = new List<string>();
                    foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(ws.Name);
                    }

                    var serializer = new JavaScriptSerializer();
                    Response.Write(serializer.Serialize(new { success = true, sheets = sheetNames }));
                }
            }
            catch (Exception ex)
            {
                ReturnError("Failed to parse sheet names: " + ex.Message);
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

            string tableName = ConfigurationManager.AppSettings["DatabaseTableName"];
            if (string.IsNullOrEmpty(tableName))
            {
                tableName = "EMPLOYEE";
            }
            tableName = tableName.Trim().ToUpper();

            string pkColumn = ConfigurationManager.AppSettings["DatabasePrimaryKeyColumn"];
            if (string.IsNullOrEmpty(pkColumn))
            {
                pkColumn = "PIS";
            }
            pkColumn = pkColumn.Trim().ToUpper();

            // Retrieve column choices from request
            string columnsToInject = Request.Form["columns"];
            if (string.IsNullOrEmpty(columnsToInject))
            {
                columnsToInject = "ACCOUNT_NUMBER"; // Fallback to account number
            }

            string[] selectedCols = columnsToInject.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            var allowedColumnsMap = GetMappableColumnsConfig();
            List<string> allowedColumns = new List<string>(allowedColumnsMap.Keys);
            List<string> colsToQuery = new List<string>();

            foreach (string col in selectedCols)
            {
                string upperCol = col.Trim().ToUpper();
                if (allowedColumns.Contains(upperCol) && !colsToQuery.Contains(upperCol))
                {
                    colsToQuery.Add(upperCol);
                }
            }

            if (colsToQuery.Count == 0)
            {
                colsToQuery.Add("ACCOUNT_NUMBER");
            }

            string connStr = ConfigurationManager.ConnectionStrings["OracleConn"].ConnectionString;
            
            // Nested mapping: employeeData[pis_number][column_name] = value
            Dictionary<string, Dictionary<string, string>> employeeData = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Fetch dynamic columns in a single query safely
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();
                    string fetchSql = string.Format("SELECT {0}, {1} FROM {2}", pkColumn, string.Join(", ", colsToQuery), tableName);
                    using (OracleCommand cmd = new OracleCommand(fetchSql, conn))
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string pisVal = CleanPisValue(reader[pkColumn].ToString());
                            if (!string.IsNullOrEmpty(pisVal))
                            {
                                var rowMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                foreach (string col in colsToQuery)
                                {
                                    object valObj = reader[col];
                                    string valStr = "";
                                    if (valObj != null && valObj != DBNull.Value)
                                    {
                                        valStr = valObj.ToString().Trim();
                                    }
                                    rowMap[col] = valStr;
                                }
                                employeeData[pisVal] = rowMap;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ReturnError("Database Error: " + ex.Message);
                return;
            }

            if (file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                ProcessCsvFile(file, colsToQuery, employeeData, allowedColumnsMap, pkColumn);
                return;
            }

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
                // Process the Excel spreadsheet in memory using EPPlus
                using (ExcelPackage package = new ExcelPackage(file.InputStream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        ReturnError("Excel file does not contain any worksheets.");
                        return;
                    }

                    bool processedAny = false;

                    // Scan all worksheets in the workbook to process selected sheets
                    foreach (ExcelWorksheet ws in package.Workbook.Worksheets)
                    {
                        // Filter by sheet name if list is provided
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
                        int pisColIndex = -1;
                        int headerRowIndex = -1;

                        for (int r = 1; r <= maxHeaderScanRows; r++)
                        {
                            for (int c = 1; c <= cCount; c++)
                            {
                                string cellText = ws.Cells[r, c].Text;
                                if (IsPisHeader(cellText, pkColumn))
                                {
                                    pisColIndex = c;
                                    headerRowIndex = r;
                                    break;
                                }
                            }
                            if (pisColIndex != -1)
                            {
                                break;
                            }
                        }

                        if (pisColIndex == -1)
                        {
                            if (selectedSheetsList.Contains(ws.Name))
                            {
                                ReturnError(string.Format("Selected sheet '{0}' does not contain a PIS column.", ws.Name));
                                return;
                            }
                            continue;
                        }

                        processedAny = true;

                        // Map column headers to append
                        Dictionary<string, int> targetColIndexes = new Dictionary<string, int>();
                        int currentColCount = cCount;
                        foreach (string col in colsToQuery)
                        {
                            currentColCount++;
                            targetColIndexes[col] = currentColCount;

                            string displayName;
                            if (!allowedColumnsMap.TryGetValue(col, out displayName))
                            {
                                displayName = col;
                            }

                            ws.Cells[headerRowIndex, currentColCount].Value = displayName;
                        }
                        
                        // Match cells by PIS values, starting directly after the header row
                        for (int r = headerRowIndex + 1; r <= rCount; r++)
                        {
                            string pis = ws.Cells[r, pisColIndex].Text;
                            if (pis != null)
                            {
                                pis = CleanPisValue(pis);
                                if (!string.IsNullOrEmpty(pis))
                                {
                                    Dictionary<string, string> record;
                                    if (employeeData.TryGetValue(pis, out record))
                                    {
                                        foreach (string col in colsToQuery)
                                        {
                                            int targetCol = targetColIndexes[col];
                                            ws.Cells[r, targetCol].Value = record[col];
                                        }
                                    }
                                    else
                                    {
                                        foreach (string col in colsToQuery)
                                        {
                                            int targetCol = targetColIndexes[col];
                                            ws.Cells[r, targetCol].Value = "Not Found";
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (!processedAny)
                    {
                        ReturnError("No valid worksheets containing a PIS column were processed.");
                        return;
                    }

                    // Output modified spreadsheet directly to browser response stream
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

        private void ReturnError(string message)
        {
            Response.Clear();
            Response.ContentType = "application/json";
            Response.StatusCode = 200; // Allow client to read structured JSON
            var serializer = new JavaScriptSerializer();
            Response.Write(serializer.Serialize(new { error = message }));
        }

        private static bool IsPisHeader(string text, string pkColumn)
        {
            if (string.IsNullOrEmpty(text)) return false;
            
            string normalized = text.Replace('\u00A0', ' ');
            normalized = normalized.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
            while (normalized.Contains("  "))
            {
                normalized = normalized.Replace("  ", " ");
            }
            
            normalized = normalized.Trim().ToUpper();
            
            if (normalized == "PIS" || 
                normalized == "PIS NO" || 
                normalized == "PIS NO." || 
                normalized == "PIS_NO" || 
                normalized == "PISNUMBER" || 
                normalized == "PIS NUMBER" || 
                normalized == "PIS_NUMBER" || 
                normalized == "PISNUM" || 
                normalized == "PIS_NUM")
            {
                return true;
            }
            
            if (!string.IsNullOrEmpty(pkColumn))
            {
                string pkNormalized = pkColumn.Trim().ToUpper();
                if (normalized == pkNormalized)
                {
                    return true;
                }
            }
            
            return false;
        }

        private static string CleanPisValue(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string cleaned = text.Replace('\u00A0', ' ');
            cleaned = cleaned.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
            while (cleaned.Contains("  "))
            {
                cleaned = cleaned.Replace("  ", " ");
            }
            return cleaned.Trim();
        }

        private void ProcessCsvFile(HttpPostedFile file, List<string> colsToQuery, Dictionary<string, Dictionary<string, string>> employeeData, Dictionary<string, string> allowedColumnsMap, string pkColumn)
        {
            try
            {
                List<List<string>> csvRows = ParseCsv(file.InputStream);
                if (csvRows.Count == 0)
                {
                    ReturnError("Uploaded CSV file is empty.");
                    return;
                }

                int pisColIndex = -1;
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
                        if (IsPisHeader(cellText, pkColumn))
                        {
                            pisColIndex = c;
                            headerRowIndex = r;
                            colCount = row.Count;
                            break;
                        }
                    }
                    if (pisColIndex != -1) break;
                }

                if (pisColIndex == -1)
                {
                    ReturnError("Could not find a 'PIS' or 'PIS NO' column in the uploaded CSV file.");
                    return;
                }

                // Map column headers to append
                var headerRow = csvRows[headerRowIndex];
                foreach (string col in colsToQuery)
                {
                    string displayName;
                    if (!allowedColumnsMap.TryGetValue(col, out displayName))
                    {
                        displayName = col;
                    }
                    headerRow.Add(displayName);
                }
                
                // Match and append values
                for (int r = headerRowIndex + 1; r < rowCount; r++)
                {
                    var row = csvRows[r];
                    while (row.Count < colCount)
                    {
                        row.Add("");
                    }

                    string pisVal = "";
                    if (pisColIndex < row.Count)
                    {
                        pisVal = CleanPisValue(row[pisColIndex]);
                    }

                    Dictionary<string, string> record = null;
                    bool hasRecord = false;
                    if (!string.IsNullOrEmpty(pisVal))
                    {
                        hasRecord = employeeData.TryGetValue(pisVal, out record);
                    }

                    foreach (string col in colsToQuery)
                    {
                        string val = hasRecord ? record[col] : "Not Found";
                        row.Add(val);
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
                            if (nextCh == '"') // Escaped quote
                            {
                                currentField.Append('"');
                                reader.Read(); // Consume the second quote
                            }
                            else
                            {
                                inQuotes = false; // End of quoted field
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
    }
}
