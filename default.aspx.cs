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
                            string pisVal = reader[pkColumn].ToString().Trim();
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

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    var dimension = worksheet.Dimension;
                    if (dimension == null)
                    {
                        ReturnError("Excel worksheet is empty.");
                        return;
                    }

                    int rowCount = dimension.End.Row;
                    int colCount = dimension.End.Column;

                    // Scan header row to identify PIS column
                    int pisColIndex = -1;
                    for (int c = 1; c <= colCount; c++)
                    {
                        string headerText = worksheet.Cells[1, c].Text;
                        if (headerText != null)
                        {
                            headerText = headerText.Trim();
                            if (string.Equals(headerText, "PIS NO", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(headerText, "PIS", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(headerText, "PIS_NO", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(headerText, pkColumn, StringComparison.OrdinalIgnoreCase))
                            {
                                pisColIndex = c;
                                break;
                            }
                        }
                    }

                    if (pisColIndex == -1)
                    {
                        ReturnError("Could not find a 'PIS NO' or 'PIS' column in the Excel header row.");
                        return;
                    }

                    // Map column headers to append
                    Dictionary<string, int> targetColIndexes = new Dictionary<string, int>();
                    foreach (string col in colsToQuery)
                    {
                        colCount++;
                        targetColIndexes[col] = colCount;

                        string displayName;
                        if (!allowedColumnsMap.TryGetValue(col, out displayName))
                        {
                            displayName = col;
                        }

                        worksheet.Cells[1, colCount].Value = displayName;
                    }
                    
                    // Match cells by PIS values
                    for (int r = 2; r <= rowCount; r++)
                    {
                        string pis = worksheet.Cells[r, pisColIndex].Text;
                        if (pis != null)
                        {
                            pis = pis.Trim();
                            if (!string.IsNullOrEmpty(pis))
                            {
                                Dictionary<string, string> record;
                                if (employeeData.TryGetValue(pis, out record))
                                {
                                    foreach (string col in colsToQuery)
                                    {
                                        int targetCol = targetColIndexes[col];
                                        worksheet.Cells[r, targetCol].Value = record[col];
                                    }
                                }
                                else
                                {
                                    foreach (string col in colsToQuery)
                                    {
                                        int targetCol = targetColIndexes[col];
                                        worksheet.Cells[r, targetCol].Value = "Not Found";
                                    }
                                }
                            }
                        }
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
    }
}
