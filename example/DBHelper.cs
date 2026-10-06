using System;
using System.Configuration;
using System.Data;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;

namespace AttendanceApp.Utils
{
    public static class DBHelper
    {
        public static string GetCompanyDBConnection()
        {
            var cs = ConfigurationManager.ConnectionStrings["CompanyDB"];
            if (cs != null && !string.IsNullOrEmpty(cs.ConnectionString)) return cs.ConnectionString;
            return "User Id=system;Password=root;Data Source=//127.0.0.1:1521/xe;";
        }

        public static string GetAttendanceDBConnection()
        {
            var cs = ConfigurationManager.ConnectionStrings["AttendanceDB"];
            if (cs != null && !string.IsNullOrEmpty(cs.ConnectionString)) return cs.ConnectionString;
            return "User Id=system;Password=root;Data Source=//127.0.0.1:1521/xe;";
        }

        [ThreadStatic]
        private static bool _inAutoClose;
        private static DateTime _lastAutoCloseCheck = DateTime.MinValue;
        private static readonly object _syncLock = new object();

        public static void TriggerAutoCloseIfNeeded()
        {
            if ((DateTime.UtcNow - _lastAutoCloseCheck).TotalSeconds < 1800) return;
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                try { AutoCloseExpiredContracts(false); } catch { }
            });
        }

        public static void AutoCloseExpiredContracts(bool force = false)
        {
            if (_inAutoClose) return;

            // Throttle checks to run at most once every 30 minutes (1800 seconds) per app domain
            if (!force && (DateTime.UtcNow - _lastAutoCloseCheck).TotalSeconds < 1800) return;

            lock (_syncLock)
            {
                if (!force && (DateTime.UtcNow - _lastAutoCloseCheck).TotalSeconds < 1800) return;
                _lastAutoCloseCheck = DateTime.UtcNow;
            }

            _inAutoClose = true;
            try
            {
                string connStr = GetAttendanceDBConnection();
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    // Find all active contract periods whose EndDate has passed
                    List<Tuple<int, DateTime>> expiredPeriods = new List<Tuple<int, DateTime>>();
                    string selectExpiredSql = "SELECT Id, EndDate FROM ContractPeriods WHERE Status = 'Active' AND EndDate < TRUNC(SYSDATE)";
                    using (OracleCommand cmd = new OracleCommand(selectExpiredSql, conn))
                    {
                        using (OracleDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int id = Convert.ToInt32(reader["Id"]);
                                DateTime endDate = Convert.ToDateTime(reader["EndDate"]);
                                expiredPeriods.Add(Tuple.Create(id, endDate));
                            }
                        }
                    }

                    if (expiredPeriods.Count > 0)
                    {
                        foreach (var period in expiredPeriods)
                        {
                            int periodId = period.Item1;
                            DateTime endDate = period.Item2;

                            using (OracleTransaction trans = conn.BeginTransaction())
                            {
                                try
                                {
                                    // a. Close ContractPeriod
                                    string closeCPSql = "UPDATE ContractPeriods SET Status = 'Closed' WHERE Id = :Id";
                                    using (OracleCommand cmd = new OracleCommand(closeCPSql, conn))
                                    {
                                        cmd.Transaction = trans;
                                        cmd.Parameters.Add(new OracleParameter("Id", periodId));
                                        cmd.ExecuteNonQuery();
                                    }

                                    // b. Find active employee engagements under this period
                                    List<Tuple<int, string>> activeEngs = new List<Tuple<int, string>>();
                                    string activeEngsSql = "SELECT Id, EmpID FROM EmployeeEngagements WHERE ContractPeriodId = :PeriodId AND EndDate IS NULL";
                                    using (OracleCommand cmd = new OracleCommand(activeEngsSql, conn))
                                    {
                                        cmd.Transaction = trans;
                                        cmd.Parameters.Add(new OracleParameter("PeriodId", periodId));
                                        using (OracleDataReader reader = cmd.ExecuteReader())
                                        {
                                            while (reader.Read())
                                            {
                                                int engId = Convert.ToInt32(reader["Id"]);
                                                string empId = reader["EmpID"].ToString();
                                                activeEngs.Add(Tuple.Create(engId, empId));
                                            }
                                        }
                                    }

                                    foreach (var eng in activeEngs)
                                    {
                                        int engId = eng.Item1;
                                        string empId = eng.Item2;

                                        // Close engagement
                                        string closeEngSql = "UPDATE EmployeeEngagements SET EndDate = :EndDate, EndReason = 'ContractEnd' WHERE Id = :Id";
                                        using (OracleCommand cmd = new OracleCommand(closeEngSql, conn))
                                        {
                                            cmd.Transaction = trans;
                                            cmd.Parameters.Add(new OracleParameter("EndDate", endDate));
                                            cmd.Parameters.Add(new OracleParameter("Id", engId));
                                            cmd.ExecuteNonQuery();
                                        }

                                        // Update Employee Master
                                        string updateEmpSql = "UPDATE Employees SET CurrentEngagementId = NULL, ContractEndDate = :ContractEndDate, Status = 'ContractEnded' WHERE MasterId = :MasterId AND CurrentEngagementId = :Id";
                                        using (OracleCommand cmd = new OracleCommand(updateEmpSql, conn))
                                        {
                                            cmd.Transaction = trans;
                                            cmd.Parameters.Add(new OracleParameter("ContractEndDate", endDate));
                                            cmd.Parameters.Add(new OracleParameter("MasterId", empId));
                                            cmd.Parameters.Add(new OracleParameter("Id", engId));
                                            cmd.ExecuteNonQuery();
                                        }
                                    }

                                    trans.Commit();
                                }
                                catch (Exception ex)
                                {
                                    trans.Rollback();
                                    System.Diagnostics.Debug.WriteLine("Error auto-closing contract period: " + ex.Message);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in AutoCloseExpiredContracts: " + ex.Message);
            }
            finally
            {
                _inAutoClose = false;
            }
        }

        private static bool IsTransientError(Exception ex)
        {
            if (ex == null) return false;
            string msg = ex.Message;
            if (msg.Contains("ORA-00904") || // invalid identifier / column missing
                msg.Contains("ORA-00942") || // table or view does not exist
                msg.Contains("ORA-00001") || // unique constraint violated
                msg.Contains("ORA-02291") || // parent key not found
                msg.Contains("ORA-02292") || // child record found
                msg.Contains("ORA-01400") || // cannot insert NULL
                msg.Contains("ORA-00936") || // missing expression
                msg.Contains("ORA-00933"))   // SQL command not properly ended
            {
                return false;
            }
            return true;
        }

        private static T RunWithRetry<T>(Func<T> operation, int maxRetries = 3, int delayMs = 500)
        {
            int attempts = 0;
            while (true)
            {
                try
                {
                    attempts++;
                    return operation();
                }
                catch (Exception ex)
                {
                    if (!IsTransientError(ex) || attempts >= maxRetries)
                    {
                        throw;
                    }
                    System.Diagnostics.Debug.WriteLine(string.Format("Transient database operation failed. Attempt {0} of {1}. Retrying in {2}ms... Error: {3}", attempts, maxRetries, delayMs * attempts, ex.Message));
                    // Wait before retrying (exponential backoff)
                    System.Threading.Thread.Sleep(delayMs * attempts);
                }
            }
        }

        private static OracleParameter[] CloneParameters(OracleParameter[] parameters)
        {
            if (parameters == null) return null;
            OracleParameter[] cloned = new OracleParameter[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
            {
                cloned[i] = new OracleParameter(parameters[i].ParameterName, parameters[i].Value)
                {
                    DbType = parameters[i].DbType,
                    Direction = parameters[i].Direction,
                    IsNullable = parameters[i].IsNullable,
                    Size = parameters[i].Size,
                    SourceColumn = parameters[i].SourceColumn,
                    SourceVersion = parameters[i].SourceVersion
                };
            }
            return cloned;
        }

        public static DataTable ExecuteQuery(string connectionString, string query, params OracleParameter[] parameters)
        {
            if (connectionString == GetAttendanceDBConnection())
            {
                EnsureSchema();
                TriggerAutoCloseIfNeeded();
            }
            return RunWithRetry(() =>
            {
                DataTable dt = new DataTable();
                using (OracleConnection conn = new OracleConnection(connectionString))
                {
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        cmd.BindByName = true;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(CloneParameters(parameters));
                        }
                        using (OracleDataAdapter sda = new OracleDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
                return dt;
            });
        }

        public static int ExecuteNonQuery(string connectionString, string query, params OracleParameter[] parameters)
        {
            if (connectionString == GetAttendanceDBConnection())
            {
                EnsureSchema();
                TriggerAutoCloseIfNeeded();
            }
            return RunWithRetry(() =>
            {
                int rowsAffected = 0;
                using (OracleConnection conn = new OracleConnection(connectionString))
                {
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        cmd.BindByName = true;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(CloneParameters(parameters));
                        }
                        conn.Open();
                        rowsAffected = cmd.ExecuteNonQuery();
                    }
                }
                return rowsAffected;
            });
        }

        public static object ExecuteScalar(string connectionString, string query, params OracleParameter[] parameters)
        {
            if (connectionString == GetAttendanceDBConnection())
            {
                EnsureSchema();
                TriggerAutoCloseIfNeeded();
            }
            return RunWithRetry(() =>
            {
                object result = null;
                using (OracleConnection conn = new OracleConnection(connectionString))
                {
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        cmd.BindByName = true;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(CloneParameters(parameters));
                        }
                        conn.Open();
                        result = cmd.ExecuteScalar();
                    }
                }
                return result;
            });
        }
        public static DataTable GetVisibleTiersDataTable(string pcno, int role)
        {
            string roleMode = System.Web.HttpContext.Current != null && System.Web.HttpContext.Current.Session != null 
                              ? (System.Web.HttpContext.Current.Session["RoleMode"]?.ToString() ?? "") 
                              : "";
            string cacheKey = string.Format("_ReqCache_TiersDt_{0}_{1}_{2}", pcno ?? "", role, roleMode);
            if (System.Web.HttpContext.Current != null)
            {
                DataTable cachedDt = System.Web.HttpContext.Current.Items[cacheKey] as DataTable;
                if (cachedDt != null) return cachedDt;
            }

            string sql;
            OracleParameter[] parameters;

            if (role == 4) // Super Admin: All Tiers
            {
                sql = @"
                    SELECT t.Id AS TierId, 
                           mc.Name || ' › ' || t.TierName || NVL2(t.RoleLabel, ' (#' || t.RoleLabel || ')', '') AS DisplayName
                    FROM Tiers t
                    JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                    ORDER BY mc.Name ASC, t.SortOrder ASC, t.TierName ASC";
                parameters = new OracleParameter[0];
            }
            else if (role == 1) // Admin: Scoped by RoleMode
            {
                if (roleMode == "PrimaryAdmin")
                {
                    sql = @"
                        SELECT t.Id AS TierId, 
                               mc.Name || ' › ' || t.TierName || NVL2(t.RoleLabel, ' (#' || t.RoleLabel || ')', '') AS DisplayName
                        FROM Tiers t
                        JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                        WHERE mc.AdminPCNO = :PCNO
                        ORDER BY mc.Name ASC, t.SortOrder ASC, t.TierName ASC";
                }
                else if (roleMode == "SecondaryAdmin")
                {
                    sql = @"
                        SELECT t.Id AS TierId, 
                               mc.Name || ' › ' || t.TierName || NVL2(t.RoleLabel, ' (#' || t.RoleLabel || ')', '') AS DisplayName
                        FROM Tiers t
                        JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                        WHERE mc.Id IN (
                               SELECT sg.MainCategoryId 
                               FROM CategoryShareGrant sg 
                               WHERE sg.SharedWithPCNO = :PCNO AND sg.IsActive = 1 AND sg.TierId IS NULL
                           )
                           OR t.Id IN (
                               SELECT sg.TierId 
                               FROM CategoryShareGrant sg 
                               WHERE sg.SharedWithPCNO = :PCNO AND sg.IsActive = 1 AND sg.TierId IS NOT NULL
                           )
                        ORDER BY mc.Name ASC, t.SortOrder ASC, t.TierName ASC";
                }
                else
                {
                    sql = @"
                        SELECT t.Id AS TierId, 
                               mc.Name || ' › ' || t.TierName || NVL2(t.RoleLabel, ' (#' || t.RoleLabel || ')', '') AS DisplayName
                        FROM Tiers t
                        JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                        WHERE mc.AdminPCNO = :PCNO
                           OR mc.Id IN (
                               SELECT sg.MainCategoryId 
                               FROM CategoryShareGrant sg 
                               WHERE sg.SharedWithPCNO = :PCNO AND sg.IsActive = 1 AND sg.TierId IS NULL
                           )
                           OR t.Id IN (
                               SELECT sg.TierId 
                               FROM CategoryShareGrant sg 
                               WHERE sg.SharedWithPCNO = :PCNO AND sg.IsActive = 1 AND sg.TierId IS NOT NULL
                           )
                        ORDER BY mc.Name ASC, t.SortOrder ASC, t.TierName ASC";
                }

                parameters = new OracleParameter[] {
                    new OracleParameter("PCNO", pcno)
                };
            }
            else // Regular User: Explicitly assigned Tiers
            {
                sql = @"
                    SELECT t.Id AS TierId, 
                           mc.Name || ' › ' || t.TierName || NVL2(t.RoleLabel, ' (#' || t.RoleLabel || ')', '') AS DisplayName
                    FROM Tiers t
                    JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                    JOIN UserTiers ut ON t.Id = ut.TierId
                    WHERE ut.PCNO = :PCNO
                    ORDER BY mc.Name ASC, t.SortOrder ASC, t.TierName ASC";
                parameters = new OracleParameter[] {
                    new OracleParameter("PCNO", pcno)
                };
            }

            DataTable dt = ExecuteQuery(GetAttendanceDBConnection(), sql, parameters);
            if (System.Web.HttpContext.Current != null && dt != null)
            {
                System.Web.HttpContext.Current.Items[cacheKey] = dt;
            }
            return dt;
        }

        public static List<int> GetVisibleTierIds(string pcno, int role)
        {
            string roleMode = System.Web.HttpContext.Current != null && System.Web.HttpContext.Current.Session != null 
                              ? (System.Web.HttpContext.Current.Session["RoleMode"]?.ToString() ?? "") 
                              : "";
            string cacheKey = string.Format("_ReqCache_TierIds_{0}_{1}_{2}", pcno ?? "", role, roleMode);
            if (System.Web.HttpContext.Current != null)
            {
                List<int> cachedList = System.Web.HttpContext.Current.Items[cacheKey] as List<int>;
                if (cachedList != null) return cachedList;
            }

            List<int> list = new List<int>();
            DataTable dt = GetVisibleTiersDataTable(pcno, role);
            foreach (DataRow row in dt.Rows)
            {
                list.Add(Convert.ToInt32(row["TierId"]));
            }
            if (System.Web.HttpContext.Current != null)
            {
                System.Web.HttpContext.Current.Items[cacheKey] = list;
            }
            return list;
        }

        private static DateTime _lastDivSyncTime = DateTime.MinValue;
        private static readonly object _divSyncLock = new object();

        public static void SyncCompanyDivisions(bool force = false)
        {
            if (!force && (DateTime.Now - _lastDivSyncTime).TotalMinutes < 15) return;

            lock (_divSyncLock)
            {
                if (!force && (DateTime.Now - _lastDivSyncTime).TotalMinutes < 15) return;

                try
                {
                    string attConnStr = GetAttendanceDBConnection();
                    using (OracleConnection conn = new OracleConnection(attConnStr))
                    {
                        conn.Open();
                        EnsureColumnExists(conn, "UserDivisions", "DivId", "ALTER TABLE UserDivisions ADD DivId NUMBER");
                        EnsureColumnExists(conn, "Employees", "DivId", "ALTER TABLE Employees ADD DivId NUMBER");
                        EnsureColumnExists(conn, "EmployeeEngagements", "DivId", "ALTER TABLE EmployeeEngagements ADD DivId NUMBER");
                    }

                    // 1. Fetch live division details from hrdata.empdetails safely
                    DataTable dtHR = null;
                    try
                    {
                        // Safely probe table schema using WHERE 1=0 to detect column availability without exceptions
                        DataTable dtSchema = new DataTable();
                        using (OracleConnection hrConn = new OracleConnection(GetCompanyDBConnection()))
                        using (OracleCommand schemaCmd = new OracleCommand("SELECT * FROM hrdata.empdetails WHERE 1=0", hrConn))
                        using (OracleDataAdapter sda = new OracleDataAdapter(schemaCmd))
                        {
                            sda.Fill(dtSchema);
                        }

                        bool hasDivId = dtSchema.Columns.Contains("DIVID") || dtSchema.Columns.Contains("divid") || dtSchema.Columns.Contains("Divid");
                        bool hasDivStatus = dtSchema.Columns.Contains("DIVSTATUS") || dtSchema.Columns.Contains("divstatus") || dtSchema.Columns.Contains("DivStatus");

                        string selectCols = hasDivId ? "divid AS id, " : "";
                        string statusFilter = hasDivStatus ? " AND divstatus = 'Y'" : "";
                        string hrSql = string.Format(@"
                            SELECT DISTINCT
                                {0}
                                CASE
                                    WHEN INSTR(divname, '/', 1, 1) <> 0
                                    THEN SUBSTR(divname, 1, INSTR(divname, '/', 1, 1) - 1)
                                    ELSE divname
                                END AS name
                            FROM hrdata.empdetails 
                            WHERE divname IS NOT NULL
                              AND divname != '*'
                              {1}
                            ORDER BY name ASC", selectCols, statusFilter);

                        dtHR = ExecuteQuery(GetCompanyDBConnection(), hrSql);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("SyncCompanyDivisions probe error: " + ex.Message);
                        dtHR = null;
                    }

                    if (dtHR == null || dtHR.Rows.Count == 0)
                    {
                        _lastDivSyncTime = DateTime.Now;
                        return;
                    }

                    if (!dtHR.Columns.Contains("Id") && !dtHR.Columns.Contains("id"))
                    {
                        dtHR.Columns.Add("id", typeof(int));
                    }

                    // Ensure local Divisions lookup table exists
                    EnsureDivisionsTableExists();

                    // Build a unified dictionary of unique division names and IDs
                    Dictionary<string, int> divMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int autoIdCounter = 1;

                    // First pass: gather existing Divisions from AttendanceDB to preserve assigned IDs
                    try
                    {
                        DataTable dtLocal = ExecuteQuery(attConnStr, "SELECT Id, Name FROM Divisions");
                        if (dtLocal != null)
                        {
                            foreach (DataRow r in dtLocal.Rows)
                            {
                                if (r["Id"] != DBNull.Value && r["Name"] != DBNull.Value)
                                {
                                    int lId = Convert.ToInt32(r["Id"]);
                                    string lName = r["Name"].ToString().Trim();
                                    if (!string.IsNullOrEmpty(lName) && !divMap.ContainsKey(lName))
                                    {
                                        divMap[lName] = lId;
                                        if (lId >= autoIdCounter) autoIdCounter = lId + 1;
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    // Check if any NULL DivId exists before running heavy backfill queries
                    bool needsUserDivsBackfill = false;
                    bool needsEmpBackfill = false;
                    bool needsEngBackfill = false;
                    try
                    {
                        object cntUD = ExecuteScalar(attConnStr, "SELECT COUNT(*) FROM UserDivisions WHERE DivId IS NULL AND ROWNUM <= 1");
                        needsUserDivsBackfill = cntUD != null && Convert.ToInt32(cntUD) > 0;
                        object cntEmp = ExecuteScalar(attConnStr, "SELECT COUNT(*) FROM Employees WHERE DivId IS NULL AND ROWNUM <= 1");
                        needsEmpBackfill = cntEmp != null && Convert.ToInt32(cntEmp) > 0;
                        object cntEng = ExecuteScalar(attConnStr, "SELECT COUNT(*) FROM EmployeeEngagements WHERE DivId IS NULL AND ROWNUM <= 1");
                        needsEngBackfill = cntEng != null && Convert.ToInt32(cntEng) > 0;
                    }
                    catch { }

                    // Second pass: process HR divisions
                    foreach (DataRow r in dtHR.Rows)
                    {
                        string colName = dtHR.Columns.Contains("name") ? "name" : "Name";
                        if (r[colName] == DBNull.Value) continue;
                        string name = r[colName].ToString().Trim();
                        if (string.IsNullOrEmpty(name) || name == "*") continue;

                        string colId = dtHR.Columns.Contains("id") ? "id" : (dtHR.Columns.Contains("Id") ? "Id" : "");
                        int divId = 0;
                        if (!string.IsNullOrEmpty(colId) && r[colId] != DBNull.Value)
                        {
                            int.TryParse(r[colId].ToString().Trim(), out divId);
                        }

                        int assignedId = divId;
                        if (assignedId <= 0)
                        {
                            if (divMap.ContainsKey(name))
                            {
                                assignedId = divMap[name];
                            }
                            else
                            {
                                assignedId = autoIdCounter++;
                            }
                        }
                        divMap[name] = assignedId;

                        string mergeSql = @"
                            MERGE INTO Divisions d
                            USING (SELECT :Id AS Id, :Name AS Name FROM DUAL) s
                            ON (d.Id = s.Id)
                            WHEN MATCHED THEN
                              UPDATE SET d.Name = s.Name WHERE d.Name <> s.Name
                            WHEN NOT MATCHED THEN
                              INSERT (Id, Name) VALUES (s.Id, s.Name)";

                        try
                        {
                            ExecuteNonQuery(attConnStr, mergeSql,
                                new OracleParameter("Id", assignedId),
                                new OracleParameter("Name", name));
                        }
                        catch
                        {
                            try
                            {
                                string insertFallback = "INSERT INTO Divisions (Id, Name) VALUES (:Id, :Name)";
                                ExecuteNonQuery(attConnStr, insertFallback,
                                    new OracleParameter("Id", assignedId),
                                    new OracleParameter("Name", name));
                            }
                            catch { }
                        }

                        // Backfill DivId for existing records ONLY if unassigned NULL DivIds were found
                        if (needsUserDivsBackfill)
                        {
                            string backfillUserDivs = "UPDATE UserDivisions SET DivId = :DivId WHERE DivId IS NULL AND (UPPER(DivisionName) = UPPER(:Name) OR UPPER(DivisionName) LIKE UPPER(:NamePrefix))";
                            try { ExecuteNonQuery(attConnStr, backfillUserDivs, new OracleParameter("DivId", assignedId), new OracleParameter("Name", name), new OracleParameter("NamePrefix", name + "/%")); } catch { }
                        }

                        if (needsEmpBackfill)
                        {
                            string backfillEmployees = "UPDATE Employees SET DivId = :DivId WHERE DivId IS NULL AND (UPPER(Department) = UPPER(:Name) OR UPPER(Department) LIKE UPPER(:NamePrefix))";
                            try { ExecuteNonQuery(attConnStr, backfillEmployees, new OracleParameter("DivId", assignedId), new OracleParameter("Name", name), new OracleParameter("NamePrefix", name + "/%")); } catch { }
                        }

                        if (needsEngBackfill)
                        {
                            string backfillEngagements = "UPDATE EmployeeEngagements SET DivId = :DivId WHERE DivId IS NULL AND (UPPER(Department) = UPPER(:Name) OR UPPER(Department) LIKE UPPER(:NamePrefix))";
                            try { ExecuteNonQuery(attConnStr, backfillEngagements, new OracleParameter("DivId", assignedId), new OracleParameter("Name", name), new OracleParameter("NamePrefix", name + "/%")); } catch { }
                        }
                    }

                    _lastDivSyncTime = DateTime.Now;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error in SyncCompanyDivisions: " + ex.Message);
                }
            }
        }

        public static DataTable GetCompanyDivisionsDataTable()
        {
            // Fast check: Try fetching from local AttendanceDB Divisions table first
            try
            {
                EnsureDivisionsTableExists();
                string sql = "SELECT Id, Name FROM Divisions ORDER BY Name ASC";
                DataTable dtLocal = ExecuteQuery(GetAttendanceDBConnection(), sql);
                if (dtLocal != null && dtLocal.Rows.Count > 0)
                {
                    // Trigger sync asynchronously/throttled in background without blocking current page load if 15 mins elapsed
                    if ((DateTime.Now - _lastDivSyncTime).TotalMinutes >= 15)
                    {
                        System.Threading.ThreadPool.QueueUserWorkItem(_ => { try { SyncCompanyDivisions(); } catch { } });
                    }
                    return dtLocal;
                }
            }
            catch { }

            // If local table is completely empty, run synchronous sync once
            SyncCompanyDivisions(force: true);

            try
            {
                EnsureDivisionsTableExists();
                string sql = "SELECT Id, Name FROM Divisions ORDER BY Name ASC";
                DataTable dtLocal = ExecuteQuery(GetAttendanceDBConnection(), sql);
                if (dtLocal != null && dtLocal.Rows.Count > 0)
                    return dtLocal;
            }
            catch { }

            // Fallback: read directly from Company HR database
            try
            {
                string sql = "SELECT DISTINCT DIVNAME AS Name FROM hrdata.empdetails WHERE DIVNAME IS NOT NULL AND DIVNAME != '*' ORDER BY DIVNAME ASC";
                DataTable dtCompany = ExecuteQuery(GetCompanyDBConnection(), sql);
                if (dtCompany != null && dtCompany.Rows.Count > 0)
                {
                    if (!dtCompany.Columns.Contains("Id"))
                    {
                        dtCompany.Columns.Add("Id", typeof(int));
                        for (int i = 0; i < dtCompany.Rows.Count; i++) dtCompany.Rows[i]["Id"] = i + 1;
                    }
                    return dtCompany;
                }
            }
            catch { }

            DataTable empty = new DataTable();
            empty.Columns.Add("Id", typeof(int));
            empty.Columns.Add("Name", typeof(string));
            return empty;
        }

        private static bool _divisionsTableEnsured = false;
        private static readonly object _divTableLock = new object();

        /// <summary>
        /// Ensures the Divisions table exists in the AttendanceDB.
        /// Creates it automatically if it is missing (e.g. oracle_setup.sql was not fully run).
        /// </summary>
        public static void EnsureDivisionsTableExists()
        {
            if (_divisionsTableEnsured) return;
            lock (_divTableLock)
            {
                if (_divisionsTableEnsured) return;

                try
                {
                    // Quick existence check
                    ExecuteScalar(GetAttendanceDBConnection(), "SELECT COUNT(*) FROM Divisions WHERE ROWNUM = 1");
                }
                catch
                {
                    // Table does not exist — create it now
                    try
                    {
                        string createTable = @"
                            CREATE TABLE Divisions (
                                Id   NUMBER        PRIMARY KEY,
                                Name VARCHAR2(100) NOT NULL UNIQUE
                            )";
                        ExecuteNonQuery(GetAttendanceDBConnection(), createTable);
                    }
                    catch { }
                }

                try
                {
                    string createSeq = "CREATE SEQUENCE SEQ_Divisions START WITH 1 INCREMENT BY 1 CACHE 20 NOCYCLE";
                    ExecuteNonQuery(GetAttendanceDBConnection(), createSeq);
                }
                catch { }

                try
                {
                    string createTrigger = @"
                        CREATE OR REPLACE TRIGGER TRG_Divisions
                        BEFORE INSERT ON Divisions
                        FOR EACH ROW
                        BEGIN
                            IF :NEW.Id IS NULL THEN
                                SELECT SEQ_Divisions.NEXTVAL INTO :NEW.Id FROM DUAL;
                            END IF;
                        END;";
                    ExecuteNonQuery(GetAttendanceDBConnection(), createTrigger);
                }
                catch { }

                _divisionsTableEnsured = true;
            }
        }

        /// <summary>
        /// Ensures a division name exists in the local Divisions table.
        /// Silently skips if the Divisions table cannot be reached.
        /// </summary>
        public static void EnsureDivisionExists(string dept)
        {
            if (string.IsNullOrWhiteSpace(dept)) return;
            try
            {
                EnsureDivisionsTableExists();
                string cleanDept = dept.Trim();
                string sqlCheck = "SELECT COUNT(*) FROM Divisions WHERE UPPER(Name) = UPPER(:Name)";
                int count = Convert.ToInt32(ExecuteScalar(GetAttendanceDBConnection(), sqlCheck, new OracleParameter("Name", cleanDept)));
                if (count == 0)
                {
                    string sqlInsert = "INSERT INTO Divisions (Name) VALUES (:Name)";
                    ExecuteNonQuery(GetAttendanceDBConnection(), sqlInsert, new OracleParameter("Name", cleanDept));
                }
            }
            catch (Exception ex)
            {
                // Non-fatal - log and continue. Employee save should not be blocked by this.
                System.Diagnostics.Debug.WriteLine("Error ensuring division exists: " + ex.Message);
            }
        }

        private static bool _schemaEnsured = false;
        private static readonly object _schemaLock = new object();

        /// <summary>
        /// Automatically checks and adds missing columns to Oracle tables if running on an older database schema.
        /// </summary>
        public static void EnsureSchema()
        {
            if (_schemaEnsured) return;
            lock (_schemaLock)
            {
                if (_schemaEnsured) return;
                try
                {
                    using (OracleConnection conn = new OracleConnection(GetAttendanceDBConnection()))
                    {
                        conn.Open();
                        EnsureColumnExists(conn, "EMPLOYEES", "TIERID", "ALTER TABLE Employees ADD (TierId NUMBER)");
                        EnsureColumnExists(conn, "EMPLOYEES", "EMPLOYEEHISTORYID", "ALTER TABLE Employees ADD (EmployeeHistoryId VARCHAR2(50))");
                        EnsureColumnExists(conn, "EMPLOYEES", "CURRENTENGAGEMENTID", "ALTER TABLE Employees ADD (CurrentEngagementId NUMBER)");
                        EnsureColumnExists(conn, "EMPLOYEEENGAGEMENTS", "TIERID", "ALTER TABLE EmployeeEngagements ADD (TierId NUMBER)");
                        EnsureColumnExists(conn, "CONTRACTPERIODS", "TIERID", "ALTER TABLE ContractPeriods ADD (TierId NUMBER)");
                        EnsureColumnExists(conn, "NOTICES", "NOTICETEXT", "ALTER TABLE Notices ADD (NoticeText NCLOB)");
                        EnsureColumnExists(conn, "NOTICES", "CATEGORY", "ALTER TABLE Notices ADD (Category VARCHAR2(100) DEFAULT 'General')");
                        EnsureColumnExists(conn, "NOTICES", "MAINCATEGORYID", "ALTER TABLE Notices ADD (MainCategoryId NUMBER NULL)");
                        EnsureColumnExists(conn, "ATTENDANCEREMARKS", "MAINCATEGORYID", "ALTER TABLE AttendanceRemarks ADD (MainCategoryId NUMBER NULL)");
                        try { using (var cmd = new OracleCommand("ALTER TABLE AttendanceRemarks MODIFY (EmpID VARCHAR2(50) NULL)", conn)) { cmd.ExecuteNonQuery(); } } catch { }
                        EnsureColumnExists(conn, "ATTPOCEDITREMARKS", "CREATEDBYROLE", "ALTER TABLE AttPocEditRemarks ADD (CreatedByRole VARCHAR2(50) DEFAULT 'POC')");
                        EnsureColumnExists(conn, "MAINCATEGORY", "ADMINPCNO", "ALTER TABLE MainCategory ADD (AdminPCNO VARCHAR2(50))");
                        EnsureColumnExists(conn, "MAINCATEGORY", "EDITDAYSALLOWED", "ALTER TABLE MainCategory ADD (EditDaysAllowed NUMBER DEFAULT 0 NOT NULL)");
                        EnsureColumnExists(conn, "MAINCATEGORY", "EDITMODE", "ALTER TABLE MainCategory ADD (EditMode NUMBER(1) DEFAULT 0 NOT NULL)");
                        try
                        {
                            string findConsSql = @"
                                SELECT uc.CONSTRAINT_NAME 
                                FROM USER_CONSTRAINTS uc 
                                JOIN USER_CONS_COLUMNS ucc ON uc.CONSTRAINT_NAME = ucc.CONSTRAINT_NAME 
                                WHERE uc.TABLE_NAME = 'MAINCATEGORY' 
                                  AND ucc.COLUMN_NAME = 'EDITMODE' 
                                  AND uc.CONSTRAINT_TYPE = 'C'";
                            using (var cmd = new OracleCommand(findConsSql, conn))
                            using (var reader = cmd.ExecuteReader())
                            {
                                var consToDrop = new List<string>();
                                while (reader.Read())
                                {
                                    consToDrop.Add(reader.GetString(0));
                                }
                                reader.Close();
                                foreach (var cName in consToDrop)
                                {
                                    try
                                    {
                                        using (var dropCmd = new OracleCommand($"ALTER TABLE MainCategory DROP CONSTRAINT {cName}", conn))
                                        {
                                            dropCmd.ExecuteNonQuery();
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }
                        EnsureColumnExists(conn, "MAINCATEGORY", "VIEWMODE", "ALTER TABLE MainCategory ADD (ViewMode NUMBER(1) DEFAULT 0 NOT NULL)");
                        EnsureColumnExists(conn, "MAINCATEGORY", "VIEWMONTHSALLOWED", "ALTER TABLE MainCategory ADD (ViewMonthsAllowed NUMBER DEFAULT 2 NOT NULL)");
                        EnsureColumnExists(conn, "MAINCATEGORY", "VIEWCUTOFFDATE", "ALTER TABLE MainCategory ADD (ViewCutoffDate DATE NULL)");
                        EnsureColumnExists(conn, "MAINCATEGORY", "VIEWAPPLIESTO", "ALTER TABLE MainCategory ADD (ViewAppliesTo NUMBER(1) DEFAULT 0 NOT NULL)");
                        EnsureColumnExists(conn, "APPUSERS", "NAME", "ALTER TABLE AppUsers ADD (Name VARCHAR2(100))");
                        EnsureTableExists(conn, "USERDIVISIONS", @"
                            CREATE TABLE UserDivisions (
                                PCNO         VARCHAR2(50)  NOT NULL,
                                DivisionName VARCHAR2(100) NOT NULL,
                                DivId        NUMBER,
                                PRIMARY KEY (PCNO, DivisionName),
                                FOREIGN KEY (PCNO) REFERENCES AppUsers(PCNO) ON DELETE CASCADE
                            )");
                        EnsureTableExists(conn, "NOTICEREADS", @"
                            CREATE TABLE NoticeReads (
                                NoticeId NUMBER NOT NULL,
                                PCNO VARCHAR2(100) NOT NULL,
                                ReadAt TIMESTAMP DEFAULT SYSTIMESTAMP NOT NULL,
                                CONSTRAINT PK_NoticeReads PRIMARY KEY (NoticeId, PCNO),
                                CONSTRAINT FK_NoticeReads_Notice FOREIGN KEY (NoticeId) REFERENCES Notices(Id) ON DELETE CASCADE
                            )");
                        EnsureTableExists(conn, "ATTENDANCEDRAFT", @"
                            CREATE TABLE AttendanceDraft (
                                Id               NUMBER        PRIMARY KEY,
                                EmpID            VARCHAR2(50)  NOT NULL,
                                EngagementId     NUMBER,
                                ContractPeriodId NUMBER,
                                Year             NUMBER(4)     NOT NULL,
                                Month            NUMBER(2)     NOT NULL,
                                Day              NUMBER(2)     NOT NULL,
                                StatusValue      NUMBER,
                                LeaveType        VARCHAR2(50),
                                IsHoliday        NUMBER(1)     DEFAULT 0 CHECK (IsHoliday IN (0, 1)),
                                AutoSat          NUMBER(1)     DEFAULT 0 CHECK (AutoSat IN (0, 1)),
                                Remarks          VARCHAR2(500),
                                EnteredByPCNO    VARCHAR2(50),
                                EnteredAt        TIMESTAMP     DEFAULT SYSTIMESTAMP NOT NULL,
                                LastEditedByPCNO VARCHAR2(50),
                                LastEditedAt     TIMESTAMP     DEFAULT SYSTIMESTAMP NOT NULL,
                                CONSTRAINT UQ_AttendanceDraft_Date UNIQUE (EmpID, Year, Month, Day),
                                FOREIGN KEY (EmpID) REFERENCES Employees(MasterId),
                                FOREIGN KEY (EngagementId) REFERENCES EmployeeEngagements(Id),
                                FOREIGN KEY (ContractPeriodId) REFERENCES ContractPeriods(Id)
                            )");
                        try
                        {
                            using (OracleCommand seqCmd = new OracleCommand("CREATE SEQUENCE SEQ_AttendanceDraft START WITH 1 INCREMENT BY 1 CACHE 20 NOCYCLE", conn))
                            {
                                seqCmd.ExecuteNonQuery();
                            }
                        }
                        catch { }
                        try
                        {
                            string trgSql = @"
                                CREATE OR REPLACE TRIGGER TRG_AttendanceDraft
                                BEFORE INSERT ON AttendanceDraft
                                FOR EACH ROW
                                BEGIN
                                    IF :NEW.Id IS NULL THEN
                                        SELECT SEQ_AttendanceDraft.NEXTVAL INTO :NEW.Id FROM DUAL;
                                    END IF;
                                END;";
                            using (OracleCommand trgCmd = new OracleCommand(trgSql, conn))
                            {
                                trgCmd.ExecuteNonQuery();
                            }
                        }
                        catch { }

                        try
                        {
                            string purgeDraftSql = @"DELETE FROM AttendanceDraft d 
                                WHERE d.IsHoliday = 1 
                                   OR (d.StatusValue IS NULL AND (d.LeaveType IS NULL OR TRIM(d.LeaveType) = '') AND (d.Remarks IS NULL OR TRIM(d.Remarks) = ''))
                                   OR EXISTS (
                                       SELECT 1 FROM Attendance a 
                                       WHERE a.EmpID = d.EmpID AND a.Year = d.Year AND a.Month = d.Month AND a.Day = d.Day 
                                         AND (a.StatusValue IS NOT NULL OR a.IsHoliday = 1 OR (a.LeaveType IS NOT NULL AND TRIM(a.LeaveType) <> ''))
                                   )";
                            using (OracleCommand purgeDraftCmd = new OracleCommand(purgeDraftSql, conn))
                            {
                                purgeDraftCmd.ExecuteNonQuery();
                            }
                        }
                        catch { }

                        EnsureTableExists(conn, "SUBUSERANCHOR", @"
                            CREATE TABLE SubUserAnchor (
                                SubUserPCNO   VARCHAR2(50)  NOT NULL,
                                AnchorPocPCNO VARCHAR2(50)  NOT NULL,
                                CreatedAt     TIMESTAMP     DEFAULT SYSTIMESTAMP NOT NULL,
                                CreatedBy     VARCHAR2(50)  NOT NULL,
                                CONSTRAINT PK_SubUserAnchor PRIMARY KEY (SubUserPCNO, AnchorPocPCNO)
                            )");

                        EnsureTableExists(conn, "WAGESALTERNATESERVICECHARGE", @"
                            CREATE TABLE WagesAlternateServiceCharge (
                                Year             NUMBER(4)     NOT NULL,
                                Month            NUMBER(2)     NOT NULL,
                                TierId           NUMBER        NOT NULL,
                                ContractPeriodId NUMBER        NOT NULL,
                                DailyRate        NUMBER(10, 2) NOT NULL,
                                ScRate           NUMBER(5, 2)  DEFAULT 3.85,
                                EpfRate          NUMBER(5, 2)  DEFAULT 13.0,
                                EpfLimit         NUMBER(10, 2) DEFAULT 15000.0,
                                EpfCappedAmount  NUMBER(10, 2) DEFAULT 1950.0,
                                ServiceCharge    NUMBER(12, 2) DEFAULT 0,
                                IsApplied        NUMBER(1)     DEFAULT 1 NOT NULL CHECK (IsApplied IN (0, 1)),
                                UpdatedBy        VARCHAR2(50),
                                UpdatedAt        TIMESTAMP     DEFAULT SYSTIMESTAMP NOT NULL,
                                PRIMARY KEY (Year, Month, TierId, ContractPeriodId),
                                FOREIGN KEY (TierId) REFERENCES Tiers(Id) ON DELETE CASCADE,
                                FOREIGN KEY (ContractPeriodId) REFERENCES ContractPeriods(Id) ON DELETE CASCADE
                            )");

                        EnsureAppUsersCompositePrimaryKey(conn);
                        EnsureAppUsersRoleConstraint(conn);
                        try {
                            using (OracleCommand nullCmd = new OracleCommand("ALTER TABLE Notices MODIFY (FilePath VARCHAR2(500) NULL)", conn)) {
                                nullCmd.ExecuteNonQuery();
                            }
                        } catch { }

                        // High-Performance Composite Indexes for 100-500+ Concurrent Users (Oracle 11g+ Compatible)
                        EnsureIndexExists(conn, "IDX_ATT_YEAR_MONTH", "CREATE INDEX IDX_ATT_YEAR_MONTH ON Attendance (Year, Month, StatusValue, EmpID)");
                        EnsureIndexExists(conn, "IDX_ATT_EMP_DATE", "CREATE INDEX IDX_ATT_EMP_DATE ON Attendance (EmpID, Year, Month)");
                        EnsureIndexExists(conn, "IDX_ATT_PERIOD", "CREATE INDEX IDX_ATT_PERIOD ON Attendance (ContractPeriodId)");
                        EnsureIndexExists(conn, "IDX_ATTDRAFT_YEAR_MONTH", "CREATE INDEX IDX_ATTDRAFT_YEAR_MONTH ON AttendanceDraft (Year, Month, EmpID)");
                        EnsureIndexExists(conn, "IDX_ATTDRAFT_EMP_DATE", "CREATE INDEX IDX_ATTDRAFT_EMP_DATE ON AttendanceDraft (EmpID, Year, Month)");
                        EnsureIndexExists(conn, "IDX_DRAFT_YM_STATUS", "CREATE INDEX IDX_DRAFT_YM_STATUS ON AttendanceDraft (Year, Month, IsHoliday, StatusValue)");
                        EnsureIndexExists(conn, "IDX_SUBUSERANCHOR_PCNO", "CREATE INDEX IDX_SUBUSERANCHOR_PCNO ON SubUserAnchor (SubUserPCNO)");
                        EnsureIndexExists(conn, "IDX_SUBUSER_POC", "CREATE INDEX IDX_SUBUSER_POC ON SubUserAnchor (AnchorPocPCNO)");
                        EnsureIndexExists(conn, "IDX_ENG_PERIOD_END", "CREATE INDEX IDX_ENG_PERIOD_END ON EmployeeEngagements (ContractPeriodId, EndDate)");
                        EnsureIndexExists(conn, "IDX_ENG_EMP_TIER", "CREATE INDEX IDX_ENG_EMP_TIER ON EmployeeEngagements (EmpID, TierId)");
                        EnsureIndexExists(conn, "IDX_ENG_DATES", "CREATE INDEX IDX_ENG_DATES ON EmployeeEngagements (StartDate, EndDate)");
                        EnsureIndexExists(conn, "IDX_EMP_TIER_STATUS", "CREATE INDEX IDX_EMP_TIER_STATUS ON Employees (TierId, Status)");
                        EnsureIndexExists(conn, "IDX_EMP_DIV_STATUS", "CREATE INDEX IDX_EMP_DIV_STATUS ON Employees (DivId, Status)");
                        EnsureIndexExists(conn, "IDX_EMP_HIST_ID", "CREATE INDEX IDX_EMP_HIST_ID ON Employees (EmployeeHistoryId)");
                        EnsureIndexExists(conn, "IDX_EMP_CURR_ENG", "CREATE INDEX IDX_EMP_CURR_ENG ON Employees (CurrentEngagementId)");
                        EnsureIndexExists(conn, "IDX_CP_TIER_STATUS", "CREATE INDEX IDX_CP_TIER_STATUS ON ContractPeriods (TierId, Status, StartDate, EndDate)");
                        EnsureIndexExists(conn, "IDX_USERTIERS_PCNO", "CREATE INDEX IDX_USERTIERS_PCNO ON UserTiers (PCNO, TierId)");
                        EnsureIndexExists(conn, "IDX_USERDIVS_PCNO", "CREATE INDEX IDX_USERDIVS_PCNO ON UserDivisions (PCNO, DivId)");
                        EnsureIndexExists(conn, "IDX_CATSHARE_SHARED", "CREATE INDEX IDX_CATSHARE_SHARED ON CategoryShareGrant (SharedWithPCNO, IsActive)");
                        EnsureIndexExists(conn, "IDX_NOTICES_CAT_DATE", "CREATE INDEX IDX_NOTICES_CAT_DATE ON Notices (Category, UploadDate)");
                        EnsureIndexExists(conn, "IDX_ATTREM_EMP_DATE", "CREATE INDEX IDX_ATTREM_EMP_DATE ON AttendanceRemarks (EmpID, RemarkDate)");
                        EnsureIndexExists(conn, "IDX_LEAVECRED_EMP", "CREATE INDEX IDX_LEAVECRED_EMP ON EmployeeLeaveCredits (EmpID, ContractPeriodId, EffectiveDate)");
                        EnsureIndexExists(conn, "IDX_POCREM_YEAR_MONTH", "CREATE INDEX IDX_POCREM_YEAR_MONTH ON AttPocEditRemarks (Year, Month, EmpID)");

                        // High-Performance Sequence Caching (Reduces dictionary write contention under 100-500 users)
                        EnsureSequenceCache(conn, "SEQ_AttendanceAuditLog", 20);
                        EnsureSequenceCache(conn, "SEQ_AttPocEditRemarks", 20);
                        EnsureSequenceCache(conn, "SEQ_EmployeeActionLogs", 20);
                        EnsureSequenceCache(conn, "SEQ_ActionLog", 20);
                        EnsureSequenceCache(conn, "SEQ_EmployeeLeaveCredits", 20);
                        EnsureSequenceCache(conn, "SEQ_AttendanceRemarks", 20);
                        EnsureSequenceCache(conn, "SEQ_AttendanceDraft", 20);
                        EnsureSequenceCache(conn, "SEQ_SubUserAnchor", 20);
                        EnsureSequenceCache(conn, "SEQ_EmployeeEngagements", 20);
                        EnsureSequenceCache(conn, "SEQ_Vendors", 20);
                        EnsureSequenceCache(conn, "SEQ_ContractPeriods", 20);
                        EnsureSequenceCache(conn, "SEQ_MainCategory", 20);
                        EnsureSequenceCache(conn, "SEQ_Tiers", 20);
                        EnsureSequenceCache(conn, "SEQ_Divisions", 20);
                    }
                    _schemaEnsured = true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("EnsureSchema error: " + ex.Message);
                }
            }
        }

        private static void EnsureAppUsersCompositePrimaryKey(OracleConnection conn)
        {
            try
            {
                string checkPkSql = @"
                    SELECT tc.CONSTRAINT_NAME, COUNT(tc.COLUMN_NAME) AS ColCount
                    FROM USER_CONSTRAINTS c
                    JOIN USER_CONS_COLUMNS tc ON c.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
                    WHERE UPPER(c.TABLE_NAME) = 'APPUSERS' AND c.CONSTRAINT_TYPE = 'P'
                    GROUP BY tc.CONSTRAINT_NAME";

                string pkName = null;
                int colCount = 0;

                using (OracleCommand cmd = new OracleCommand(checkPkSql, conn))
                {
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            pkName = reader["CONSTRAINT_NAME"].ToString();
                            colCount = Convert.ToInt32(reader["ColCount"]);
                        }
                    }
                }

                // If primary key is single-column (only PCNO), drop old single-column PK with CASCADE and create composite (PCNO, Role) PK
                if (colCount == 1 || (colCount == 0 && pkName != null))
                {
                    if (!string.IsNullOrEmpty(pkName))
                    {
                        try
                        {
                            using (OracleCommand dropCmd = new OracleCommand("ALTER TABLE AppUsers DROP CONSTRAINT " + pkName + " CASCADE", conn))
                            {
                                dropCmd.ExecuteNonQuery();
                            }
                        }
                        catch { }
                    }

                    try
                    {
                        using (OracleCommand dropPkCmd = new OracleCommand("ALTER TABLE AppUsers DROP PRIMARY KEY CASCADE", conn))
                        {
                            dropPkCmd.ExecuteNonQuery();
                        }
                    }
                    catch { }

                    try
                    {
                        using (OracleCommand dedupCmd = new OracleCommand("DELETE FROM AppUsers a WHERE ROWID < (SELECT MAX(ROWID) FROM AppUsers b WHERE a.PCNO = b.PCNO AND a.Role = b.Role)", conn))
                        {
                            dedupCmd.ExecuteNonQuery();
                        }
                    }
                    catch { }

                    try
                    {
                        using (OracleCommand addCmd = new OracleCommand("ALTER TABLE AppUsers ADD PRIMARY KEY (PCNO, Role)", conn))
                        {
                            addCmd.ExecuteNonQuery();
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("EnsureAppUsersCompositePrimaryKey error: " + ex.Message);
            }
        }

        public static void EnsureAppUsersRoleConstraint()
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(GetAttendanceDBConnection()))
                {
                    conn.Open();
                    EnsureAppUsersRoleConstraint(conn);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("EnsureAppUsersRoleConstraint error: " + ex.Message);
            }
        }

        public static void EnsureAppUsersRoleConstraint(OracleConnection conn)
        {
            try
            {
                // 1. Explicitly drop known system-generated and custom check constraints if present
                string[] knownConstraints = new string[] { "SYS_C009993", "CHK_APPUSERS_ROLE", "CHK_AppUsers_Role", "SYS_C009994", "SYS_C009995", "SYS_C009992" };
                foreach (var kName in knownConstraints)
                {
                    try
                    {
                        using (OracleCommand dropCmd = new OracleCommand("ALTER TABLE AppUsers DROP CONSTRAINT " + kName, conn))
                        {
                            dropCmd.ExecuteNonQuery();
                        }
                    }
                    catch { }
                }

                // 2. Dynamically find all Check constraints on AppUsers
                try
                {
                    string checkSql = "SELECT CONSTRAINT_NAME FROM USER_CONSTRAINTS WHERE UPPER(TABLE_NAME) = 'APPUSERS' AND CONSTRAINT_TYPE = 'C'";
                    List<string> constraintsToDrop = new List<string>();
                    using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string cName = reader["CONSTRAINT_NAME"].ToString();
                            // Target any system check constraint or role-related constraint on AppUsers
                            if (cName.StartsWith("SYS_C", StringComparison.OrdinalIgnoreCase) || 
                                cName.IndexOf("ROLE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                cName.IndexOf("APPUSERS", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                constraintsToDrop.Add(cName);
                            }
                        }
                    }

                    foreach (var cName in constraintsToDrop)
                    {
                        try
                        {
                            using (OracleCommand dropCmd = new OracleCommand("ALTER TABLE AppUsers DROP CONSTRAINT " + cName, conn))
                            {
                                dropCmd.ExecuteNonQuery();
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // 3. Ensure columns are NOT NULL
                try
                {
                    using (OracleCommand modCmd = new OracleCommand("ALTER TABLE AppUsers MODIFY (PCNO VARCHAR2(50) NOT NULL)", conn))
                    {
                        modCmd.ExecuteNonQuery();
                    }
                }
                catch { }

                try
                {
                    using (OracleCommand modCmd = new OracleCommand("ALTER TABLE AppUsers MODIFY (Role NUMBER(1) DEFAULT 0 NOT NULL)", conn))
                    {
                        modCmd.ExecuteNonQuery();
                    }
                }
                catch { }

                // 4. Add the expanded CHECK constraint allowing roles 0..7
                try
                {
                    using (OracleCommand addCmd = new OracleCommand("ALTER TABLE AppUsers ADD CONSTRAINT CHK_AppUsers_Role CHECK (Role IN (0, 1, 2, 3, 4, 5, 6, 7))", conn))
                    {
                        addCmd.ExecuteNonQuery();
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("EnsureAppUsersRoleConstraint error: " + ex.Message);
            }
        }

        private static void EnsureTableExists(OracleConnection conn, string tableName, string createTableSql)
        {
            try
            {
                string checkSql = "SELECT COUNT(*) FROM USER_TABLES WHERE UPPER(TABLE_NAME) = :TName";
                using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                {
                    cmd.Parameters.Add(new OracleParameter("TName", tableName.ToUpper()));
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count == 0)
                    {
                        using (OracleCommand createCmd = new OracleCommand(createTableSql, conn))
                        {
                            createCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("EnsureTableExists ({0}) error: {1}", tableName, ex.Message));
            }
        }

        private static void EnsureColumnExists(OracleConnection conn, string tableName, string columnName, string alterSql)
        {
            try
            {
                string checkSql = "SELECT COUNT(*) FROM USER_TAB_COLUMNS WHERE UPPER(TABLE_NAME) = :TName AND UPPER(COLUMN_NAME) = :CName";
                using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                {
                    cmd.Parameters.Add(new OracleParameter("TName", tableName.ToUpper()));
                    cmd.Parameters.Add(new OracleParameter("CName", columnName.ToUpper()));
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count == 0)
                    {
                        using (OracleCommand alterCmd = new OracleCommand(alterSql, conn))
                        {
                            alterCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("EnsureColumnExists ({0}.{1}) error: {2}", tableName, columnName, ex.Message));
            }
        }

        private static void EnsureIndexExists(OracleConnection conn, string indexName, string createIndexSql)
        {
            try
            {
                string checkSql = "SELECT COUNT(*) FROM USER_INDEXES WHERE UPPER(INDEX_NAME) = :IName";
                using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                {
                    cmd.Parameters.Add(new OracleParameter("IName", indexName.ToUpper()));
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count == 0)
                    {
                        using (OracleCommand createCmd = new OracleCommand(createIndexSql, conn))
                        {
                            createCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("ORA-01408") && !ex.Message.Contains("ORA-00955"))
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("EnsureIndexExists ({0}) error: {1}", indexName, ex.Message));
                }
            }
        }

        private static void EnsureSequenceCache(OracleConnection conn, string seqName, int cacheSize = 20)
        {
            try
            {
                string checkSql = "SELECT CACHE_SIZE FROM USER_SEQUENCES WHERE UPPER(SEQUENCE_NAME) = :SName";
                using (OracleCommand cmd = new OracleCommand(checkSql, conn))
                {
                    cmd.Parameters.Add(new OracleParameter("SName", seqName.ToUpper()));
                    object res = cmd.ExecuteScalar();
                    if (res != null && res != DBNull.Value && Convert.ToInt32(res) < cacheSize)
                    {
                        using (OracleCommand alterCmd = new OracleCommand(string.Format("ALTER SEQUENCE {0} CACHE {1}", seqName, cacheSize), conn))
                        {
                            alterCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("EnsureSequenceCache ({0}) error: {1}", seqName, ex.Message));
            }
        }

        public static void BackfillHolidaysForEmployee(string masterId)
        {
            if (string.IsNullOrWhiteSpace(masterId) || masterId.StartsWith("GLOBAL", StringComparison.OrdinalIgnoreCase)) return;

            try
            {
                string connStr = GetAttendanceDBConnection();
                string query = @"
                    INSERT INTO Attendance (EmpID, Year, Month, Day, StatusValue, IsHoliday, LeaveType, AutoSat, Remarks)
                    SELECT :EmpID, h.Year, h.Month, h.Day, NULL, 1, '', 0, h.Remarks
                    FROM (
                        SELECT Year, Month, Day, MAX(Remarks) AS Remarks
                        FROM Attendance
                        WHERE IsHoliday = 1 AND Day > 0 AND EmpID NOT LIKE 'GLOBAL%'
                        GROUP BY Year, Month, Day
                    ) h
                    WHERE EXISTS (
                        SELECT 1 FROM EmployeeEngagements ee
                        WHERE ee.EmpID = :EmpID2
                          AND TO_DATE(h.Year || '-' || (h.Month + 1) || '-' || h.Day, 'YYYY-MM-DD') >= TRUNC(ee.StartDate)
                          AND (ee.EndDate IS NULL OR TO_DATE(h.Year || '-' || (h.Month + 1) || '-' || h.Day, 'YYYY-MM-DD') <= TRUNC(ee.EndDate))
                    )
                    AND NOT EXISTS (
                        SELECT 1 FROM Attendance a 
                        WHERE a.EmpID = :EmpID3 AND a.Year = h.Year AND a.Month = h.Month AND a.Day = h.Day
                    )";

                ExecuteNonQuery(connStr, query,
                    new OracleParameter("EmpID", masterId),
                    new OracleParameter("EmpID2", masterId),
                    new OracleParameter("EmpID3", masterId));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in BackfillHolidaysForEmployee: " + ex.Message);
            }
        }

        public static void BackfillHolidaysForAllEmployees()
        {
            try
            {
                string connStr = GetAttendanceDBConnection();
                string query = @"
                    INSERT INTO Attendance (EmpID, Year, Month, Day, StatusValue, IsHoliday, LeaveType, AutoSat, Remarks)
                    SELECT e.MasterId, h.Year, h.Month, h.Day, NULL, 1, '', 0, h.Remarks
                    FROM Employees e
                    CROSS JOIN (
                        SELECT Year, Month, Day, MAX(Remarks) AS Remarks
                        FROM Attendance
                        WHERE IsHoliday = 1 AND Day > 0 AND EmpID NOT LIKE 'GLOBAL%'
                        GROUP BY Year, Month, Day
                    ) h
                    WHERE e.MasterId NOT LIKE 'GLOBAL%'
                      AND EXISTS (
                        SELECT 1 FROM EmployeeEngagements ee
                        WHERE ee.EmpID = e.MasterId
                          AND TO_DATE(h.Year || '-' || (h.Month + 1) || '-' || h.Day, 'YYYY-MM-DD') >= TRUNC(ee.StartDate)
                          AND (ee.EndDate IS NULL OR TO_DATE(h.Year || '-' || (h.Month + 1) || '-' || h.Day, 'YYYY-MM-DD') <= TRUNC(ee.EndDate))
                      )
                      AND NOT EXISTS (
                        SELECT 1 FROM Attendance a 
                        WHERE a.EmpID = e.MasterId AND a.Year = h.Year AND a.Month = h.Month AND a.Day = h.Day
                    )";

                ExecuteNonQuery(connStr, query);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in BackfillHolidaysForAllEmployees: " + ex.Message);
            }
        }

        public static List<UserRoleOption> GetAvailableUserRoles(string pcno)
        {
            var roles = new List<UserRoleOption>();
            if (string.IsNullOrEmpty(pcno)) return roles;

            try
            {
                // Fetch all active/assigned roles in AppUsers for this PCNO
                string queryRoles = "SELECT Role FROM AppUsers WHERE PCNO = :PCNO";
                DataTable dtUserRoles = ExecuteQuery(GetAttendanceDBConnection(), queryRoles, new OracleParameter("PCNO", pcno));
                List<int> userRoles = new List<int>();
                if (dtUserRoles != null)
                {
                    foreach (DataRow r in dtUserRoles.Rows)
                    {
                        if (r["Role"] != DBNull.Value) userRoles.Add(Convert.ToInt32(r["Role"]));
                    }
                }

                // 1. Super Admin
                if (userRoles.Contains(4) && !userRoles.Contains(5))
                {
                    roles.Add(new UserRoleOption
                    {
                        RoleMode = "SuperAdmin",
                        Title = "Super Administrator",
                        Subtitle = "Full system configuration & global access",
                        EffectiveRole = 4,
                        Icon = "fas fa-crown",
                        BadgeColor = "#f59e0b"
                    });
                }

                // 2. Secondary Category Admin (Category Sharing)
                string queryShared = @"
                    SELECT DISTINCT mc.Name 
                    FROM CategoryShareGrant sg 
                    JOIN MainCategory mc ON sg.MainCategoryId = mc.Id 
                    WHERE sg.SharedWithPCNO = :PCNO AND sg.IsActive = 1";
                DataTable dtShared = ExecuteQuery(GetAttendanceDBConnection(), queryShared, new OracleParameter("PCNO", pcno));

                bool isRevokedAdmin = userRoles.Contains(2);

                if (dtShared.Rows.Count > 0 && !isRevokedAdmin)
                {
                    List<string> sharedNames = new List<string>();
                    foreach (DataRow r in dtShared.Rows) sharedNames.Add(r["Name"].ToString());

                    roles.Add(new UserRoleOption
                    {
                        RoleMode = "SecondaryAdmin",
                        Title = "Secondary Admin (Shared)",
                        Subtitle = "Shared Categories (" + string.Join(", ", sharedNames) + ")",
                        EffectiveRole = 1,
                        Icon = "fas fa-share-alt",
                        BadgeColor = "#0284c7"
                    });
                }

                // 3. Primary Category Admin (Category Owner)
                string queryMC = "SELECT Name FROM MainCategory WHERE AdminPCNO = :PCNO";
                DataTable dtMC = ExecuteQuery(GetAttendanceDBConnection(), queryMC, new OracleParameter("PCNO", pcno));

                string qAnyGrants = "SELECT COUNT(*) FROM CategoryShareGrant WHERE SharedWithPCNO = :PCNO";
                object anyGrantsObj = ExecuteScalar(GetAttendanceDBConnection(), qAnyGrants, new OracleParameter("PCNO", pcno));
                bool hasEverBeenGuest = anyGrantsObj != null && Convert.ToInt32(anyGrantsObj) > 0;

                // A user is a Primary Category Admin if they own a main category, OR if they were explicitly registered as an Admin AND are not a guest category sharing user
                bool hasPrimaryAdminRole = !isRevokedAdmin && (dtMC.Rows.Count > 0 || (userRoles.Contains(1) && dtShared.Rows.Count == 0 && !hasEverBeenGuest));
                if (hasPrimaryAdminRole)
                {
                    string subtitleText = "Category Owner (Unassigned Category)";
                    if (dtMC.Rows.Count > 0)
                    {
                        List<string> mcNames = new List<string>();
                        foreach (DataRow r in dtMC.Rows) mcNames.Add(r["Name"].ToString());
                        subtitleText = "Category Owner (" + string.Join(", ", mcNames) + ")";
                    }

                    roles.Add(new UserRoleOption
                    {
                        RoleMode = "PrimaryAdmin",
                        Title = "Primary Category Admin",
                        Subtitle = subtitleText,
                        EffectiveRole = 1,
                        Icon = "fas fa-user-shield",
                        BadgeColor = "#4f46e5"
                    });
                }

                // 4. Regular User (POC) & Sub User (Data Entry)
                string queryDivs = "SELECT DivisionName FROM UserDivisions WHERE PCNO = :PCNO";
                DataTable dtDivs = ExecuteQuery(GetAttendanceDBConnection(), queryDivs, new OracleParameter("PCNO", pcno));

                string qAnchor = "SELECT COUNT(*) FROM SubUserAnchor WHERE SubUserPCNO = :PCNO";
                object anchorObj = ExecuteScalar(GetAttendanceDBConnection(), qAnchor, new OracleParameter("PCNO", pcno));
                bool isSubUserInAnchor = anchorObj != null && anchorObj != DBNull.Value && Convert.ToInt32(anchorObj) > 0;

                bool isRevokedPoc = userRoles.Contains(3);
                bool isRevokedSubUser = userRoles.Contains(7);
                bool isSubUserRegistered = userRoles.Contains(6) || userRoles.Contains(7) || isSubUserInAnchor;

                // User is POC if explicitly granted Role 0 (and not revoked with Role 3),
                // OR (legacy fallback) if they have UserDivisions and are NOT a Sub User (nor an Admin/SuperAdmin).
                bool hasPocRole = !isRevokedPoc && (userRoles.Contains(0) || (dtDivs.Rows.Count > 0 && !isSubUserRegistered && !userRoles.Contains(1) && !userRoles.Contains(2) && !userRoles.Contains(4) && !userRoles.Contains(5)));
                bool hasExplicitSubUser = !isRevokedSubUser && (userRoles.Contains(6) || isSubUserInAnchor);

                List<string> divNames = new List<string>();
                if (dtDivs != null)
                {
                    foreach (DataRow r in dtDivs.Rows) divNames.Add(r["DivisionName"].ToString());
                }
                string divText = divNames.Count > 0 ? string.Join(", ", divNames) : "All Divisions";

                if (hasPocRole)
                {
                    roles.Add(new UserRoleOption
                    {
                        RoleMode = "RegularUser",
                        Title = "Regular User (POC)",
                        Subtitle = "Directorate POC (" + divText + ")",
                        EffectiveRole = 0,
                        Icon = "fas fa-user-check",
                        BadgeColor = "#64748b"
                    });

                    // Implicit Sub User: every active POC gets Sub User mode with the exact same scope
                    if (!isRevokedSubUser)
                    {
                        roles.Add(new UserRoleOption
                        {
                            RoleMode = "SubUser",
                            Title = "Sub User (Data Entry)",
                            Subtitle = "Draft Entry (" + divText + ")",
                            EffectiveRole = 6,
                            Icon = "fas fa-user-edit",
                            BadgeColor = "#0284c7"
                        });
                    }
                }
                else if (hasExplicitSubUser)
                {
                    roles.Add(new UserRoleOption
                    {
                        RoleMode = "SubUser",
                        Title = "Sub User (Data Entry)",
                        Subtitle = "Draft Entry (" + divText + ")",
                        EffectiveRole = 6,
                        Icon = "fas fa-user-edit",
                        BadgeColor = "#0284c7"
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in GetAvailableUserRoles: " + ex.Message);
            }

            return roles;
        }

        public static bool SwitchUserRole(string pcno, string targetMode, System.Web.SessionState.HttpSessionState session)
        {
            if (string.IsNullOrEmpty(pcno) || string.IsNullOrEmpty(targetMode) || session == null) return false;

            try
            {
                var userRoles = GetAvailableUserRoles(pcno);
                var selectedRole = userRoles.Find(r => string.Equals(r.RoleMode, targetMode, StringComparison.OrdinalIgnoreCase));
                if (selectedRole == null) return false;

                session["Role"] = selectedRole.EffectiveRole;
                session["RoleMode"] = selectedRole.RoleMode;
                session["UserRoles"] = userRoles;

                // Load and refresh allowed divisions
                List<string> allowedDivisions = new List<string>();
                if (selectedRole.EffectiveRole != 1 && selectedRole.EffectiveRole != 4)
                {
                    try
                    {
                        EnsureDivisionsTableExists();
                        string queryDivs = "SELECT DivisionName FROM UserDivisions WHERE PCNO = :PCNO";
                        DataTable dtDivs = ExecuteQuery(GetAttendanceDBConnection(), queryDivs, new OracleParameter("PCNO", pcno));
                        foreach (DataRow row in dtDivs.Rows)
                        {
                            if (row["DivisionName"] != DBNull.Value && !string.IsNullOrWhiteSpace(row["DivisionName"].ToString()))
                            {
                                allowedDivisions.Add(row["DivisionName"].ToString());
                            }
                        }
                        if (allowedDivisions.Count == 0)
                        {
                            allowedDivisions.Add("D-USER");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Error fetching user divisions in SwitchUserRole: " + ex.Message);
                    }

                    session["AllowedDivisions"] = allowedDivisions;
                    session["Division"] = allowedDivisions.Count > 0 ? allowedDivisions[0] : "D-USER";
                }
                else
                {
                    session["Division"] = selectedRole.EffectiveRole == 4 ? "D-SUPERADMIN" : "D-ADMIN";
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in SwitchUserRole: " + ex.Message);
                return false;
            }
        }

        public static PocViewRestrictionInfo GetPocViewRestriction(string pcno, string targetPage = "Attendance", string categoryFilter = null)
        {
            string cacheKey = string.Format("_ReqCache_PocRestr_{0}_{1}_{2}", pcno ?? "", targetPage ?? "Attendance", categoryFilter ?? "All");
            if (System.Web.HttpContext.Current != null)
            {
                PocViewRestrictionInfo cachedInfo = System.Web.HttpContext.Current.Items[cacheKey] as PocViewRestrictionInfo;
                if (cachedInfo != null) return cachedInfo;
            }

            var info = new PocViewRestrictionInfo
            {
                IsRestricted = false,
                ViewMode = 0,
                ViewMonthsAllowed = 2,
                ViewCutoffDate = null,
                ViewAppliesTo = 0,
                MinAllowedDate = null,
                MaxAllowedDate = null,
                Description = "All Months (Unrestricted)"
            };

            if (string.IsNullOrEmpty(pcno))
                return info;

            try
            {
                EnsureSchema();

                string query = @"
                    SELECT DISTINCT mc.Id, mc.Name, 
                           COALESCE(mc.ViewMode, 0) AS ViewMode, 
                           COALESCE(mc.ViewMonthsAllowed, 2) AS ViewMonthsAllowed, 
                           mc.ViewCutoffDate, 
                           COALESCE(mc.ViewAppliesTo, 0) AS ViewAppliesTo
                    FROM UserTiers ut
                    JOIN Tiers t ON ut.TierId = t.Id
                    JOIN MainCategory mc ON t.MainCategoryId = mc.Id
                    WHERE ut.PCNO = :PCNO
                       OR ut.PCNO IN (SELECT AnchorPocPCNO FROM SubUserAnchor WHERE SubUserPCNO = :PCNO)";

                List<OracleParameter> prms = new List<OracleParameter>
                {
                    new OracleParameter("PCNO", pcno)
                };

                if (!string.IsNullOrEmpty(categoryFilter) && categoryFilter != "All")
                {
                    int catIdVal;
                    if (int.TryParse(categoryFilter, out catIdVal))
                    {
                        query += " AND (mc.Id = :CatId OR mc.Name = :CatName OR t.Id = :CatId)";
                        prms.Add(new OracleParameter("CatId", catIdVal));
                        prms.Add(new OracleParameter("CatName", categoryFilter));
                    }
                    else
                    {
                        query += " AND mc.Name = :CatName";
                        prms.Add(new OracleParameter("CatName", categoryFilter));
                    }
                }

                DataTable dt = ExecuteQuery(GetAttendanceDBConnection(), query, prms.ToArray());
                if (dt.Rows.Count == 0)
                {
                    return info;
                }

                DateTime today = DateTime.Today;
                bool anyApplicable = false;
                DateTime? globalMin = null;
                DateTime? globalMax = null;
                List<string> descriptions = new List<string>();

                foreach (DataRow row in dt.Rows)
                {
                    int vMode = Convert.ToInt32(row["ViewMode"]);
                    int vMonths = Convert.ToInt32(row["ViewMonthsAllowed"]);
                    DateTime? vCutoff = row["ViewCutoffDate"] != DBNull.Value ? Convert.ToDateTime(row["ViewCutoffDate"]) : (DateTime?)null;
                    int vAppliesTo = Convert.ToInt32(row["ViewAppliesTo"]); // 0=Both, 1=Attendance, 2=Ledger
                    int catId = Convert.ToInt32(row["Id"]);

                    // Check if this restriction applies to the requested target page
                    bool applies = false;
                    if (vAppliesTo == 0) applies = true;
                    else if (targetPage.Equals("Attendance", StringComparison.OrdinalIgnoreCase) && vAppliesTo == 1) applies = true;
                    else if (targetPage.Equals("Ledger", StringComparison.OrdinalIgnoreCase) && vAppliesTo == 2) applies = true;

                    if (!applies || vMode == 0)
                    {
                        continue;
                    }

                    anyApplicable = true;
                    DateTime? catMin = null;
                    DateTime? catMax = null;
                    string catDesc = "";

                    if (vMode == 1) // Current Month Only
                    {
                        catMin = new DateTime(today.Year, today.Month, 1);
                        catMax = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));
                        catDesc = "Current Month (" + today.ToString("MMM yyyy") + ")";
                    }
                    else if (vMode == 2) // Last N Months (Current & Previous)
                    {
                        int backCount = Math.Max(1, vMonths) - 1;
                        catMin = new DateTime(today.Year, today.Month, 1).AddMonths(-backCount);
                        catMax = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));
                        catDesc = $"Last {vMonths} Months ({catMin.Value:MMM yyyy} - {today:MMM yyyy})";
                    }
                    else if (vMode == 3) // Active Contract Period
                    {
                        string qCp = @"
                            SELECT MIN(cp.StartDate) AS MinStart, MAX(cp.EndDate) AS MaxEnd
                            FROM ContractPeriods cp
                            JOIN Tiers t ON cp.TierId = t.Id
                            WHERE t.MainCategoryId = :CatId
                              AND cp.Status = 'Active'
                              AND cp.StartDate <= :Today AND (cp.EndDate IS NULL OR cp.EndDate >= :Today)";
                        DataTable dtCp = ExecuteQuery(GetAttendanceDBConnection(), qCp, 
                            new OracleParameter("CatId", catId),
                            new OracleParameter("Today", today));

                        if (dtCp.Rows.Count > 0 && dtCp.Rows[0]["MinStart"] != DBNull.Value)
                        {
                            catMin = Convert.ToDateTime(dtCp.Rows[0]["MinStart"]);
                            catMin = new DateTime(catMin.Value.Year, catMin.Value.Month, 1);
                            if (dtCp.Rows[0]["MaxEnd"] != DBNull.Value)
                            {
                                DateTime me = Convert.ToDateTime(dtCp.Rows[0]["MaxEnd"]);
                                catMax = new DateTime(me.Year, me.Month, DateTime.DaysInMonth(me.Year, me.Month));
                            }
                            catDesc = $"Active Contract ({catMin.Value:MMM yyyy} - {(catMax.HasValue ? catMax.Value.ToString("MMM yyyy") : "Present")})";
                        }
                        else
                        {
                            catMin = new DateTime(today.Year, today.Month, 1);
                            catDesc = "Active Contract (Current Month)";
                        }
                    }
                    else if (vMode == 4) // From Specific Month Onwards (Cutoff Date)
                    {
                        if (vCutoff.HasValue)
                        {
                            catMin = new DateTime(vCutoff.Value.Year, vCutoff.Value.Month, 1);
                            catDesc = $"From {catMin.Value:MMM yyyy} Onwards";
                        }
                        else
                        {
                            catMin = new DateTime(today.Year, today.Month, 1);
                            catDesc = $"From {today:MMM yyyy} Onwards";
                        }
                    }

                    // Union boundaries: most expansive allowed range across user's categories
                    if (!globalMin.HasValue || (catMin.HasValue && catMin.Value < globalMin.Value))
                        globalMin = catMin;

                    if (catMax.HasValue)
                    {
                        if (!globalMax.HasValue || catMax.Value > globalMax.Value)
                            globalMax = catMax;
                    }

                    if (!descriptions.Contains(catDesc))
                        descriptions.Add(catDesc);
                }

                if (anyApplicable)
                {
                    info.IsRestricted = true;
                    info.MinAllowedDate = globalMin;
                    info.MaxAllowedDate = globalMax;
                    info.Description = string.Join("; ", descriptions);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in GetPocViewRestriction: " + ex.Message);
            }

            if (System.Web.HttpContext.Current != null && info != null)
            {
                System.Web.HttpContext.Current.Items[cacheKey] = info;
            }

            return info;
        }

        public static AttendanceCompletenessResult CheckMonthAttendanceCompleteness(int year, int month, string category, int? contractPeriodId, int role, string pcno)
        {
            AttendanceCompletenessResult result = new AttendanceCompletenessResult
            {
                HasIncomplete = false,
                MissingCount = 0,
                UnspecifiedZeroCount = 0,
                PendingPairingCount = 0,
                MissingExamples = new List<string>(),
                UnspecifiedZeroExamples = new List<string>(),
                PendingPairingExamples = new List<string>(),
                WarningMessage = ""
            };

            // Only relevant for Admin roles
            if (role != 1 && role != 4) return result;

            try
            {
                int dbMonth = month + 1;
                DateTime firstDay = new DateTime(year, dbMonth, 1);
                int daysInMonth = DateTime.DaysInMonth(year, dbMonth);
                DateTime lastDay = new DateTime(year, dbMonth, daysInMonth);

                int catTierId = 0;
                int.TryParse(category, out catTierId);

                List<string> targetTierIds = new List<string>();
                if (contractPeriodId.HasValue && contractPeriodId.Value > 0)
                {
                    // contractPeriodId specified; will filter by contract period
                }
                else if (catTierId > 0)
                {
                    targetTierIds.Add(catTierId.ToString());
                }
                else
                {
                    // If "All" or empty, gather all visible tiers
                    DataTable dtTiers = GetVisibleTiersDataTable(pcno, role);
                    foreach (DataRow row in dtTiers.Rows)
                    {
                        targetTierIds.Add(row["TierId"].ToString());
                    }
                }

                using (OracleConnection conn = new OracleConnection(GetAttendanceDBConnection()))
                {
                    conn.Open();

                    // 1. Fetch Employees
                    string empSql;
                    List<OracleParameter> empParams = new List<OracleParameter>();

                    if (contractPeriodId.HasValue && contractPeriodId.Value > 0)
                    {
                        empSql = @"
                            SELECT DISTINCT e.ID, e.Name, e.Department, e.JoinDate, e.ResignDate, e.LeaveBalance,
                                   cp.EndDate AS ContractEndDate, e.MasterId, ee.TierId
                            FROM Employees e
                            JOIN EmployeeEngagements ee ON e.MasterId = ee.EmpID
                            JOIN ContractPeriods cp ON ee.ContractPeriodId = cp.Id
                            WHERE ee.ContractPeriodId = :SelectedCpId";
                        empParams.Add(new OracleParameter("SelectedCpId", contractPeriodId.Value));
                    }
                    else if (targetTierIds.Count > 0)
                    {
                        List<string> pNames = new List<string>();
                        for (int i = 0; i < targetTierIds.Count; i++)
                        {
                            string pName = "VTierId" + i;
                            pNames.Add(":" + pName);
                            empParams.Add(new OracleParameter(pName, targetTierIds[i]));
                        }
                        empSql = string.Format(@"
                            SELECT DISTINCT e.ID, e.Name, e.Department, e.JoinDate, e.ResignDate, e.LeaveBalance,
                                   e.ContractEndDate, e.MasterId, e.TierId
                            FROM Employees e
                            WHERE e.TierId IN ({0})", string.Join(", ", pNames));
                    }
                    else
                    {
                        return result;
                    }

                    DataTable dtEmp = new DataTable();
                    using (OracleCommand cmdEmp = new OracleCommand(empSql, conn))
                    {
                        cmdEmp.BindByName = true;
                        foreach (var p in empParams) cmdEmp.Parameters.Add(p);
                        using (OracleDataAdapter da = new OracleDataAdapter(cmdEmp))
                        {
                            da.Fill(dtEmp);
                        }
                    }

                    if (dtEmp.Rows.Count == 0) return result;

                    // 2. Fetch Engagements
                    string engSql = @"
                        SELECT ee.EmpID, ee.StartDate, ee.EndDate, cp.EndDate AS CpEndDate
                        FROM EmployeeEngagements ee
                        JOIN ContractPeriods cp ON ee.ContractPeriodId = cp.Id
                        WHERE (ee.StartDate <= :LastDay AND (ee.EndDate IS NULL OR ee.EndDate >= :FirstDay))";
                    List<OracleParameter> engParams = new List<OracleParameter> {
                        new OracleParameter("LastDay", lastDay),
                        new OracleParameter("FirstDay", firstDay)
                    };
                    if (contractPeriodId.HasValue && contractPeriodId.Value > 0)
                    {
                        engSql += " AND ee.ContractPeriodId = :SelectedCpId";
                        engParams.Add(new OracleParameter("SelectedCpId", contractPeriodId.Value));
                    }

                    DataTable dtEng = new DataTable();
                    using (OracleCommand cmdEng = new OracleCommand(engSql, conn))
                    {
                        cmdEng.BindByName = true;
                        foreach (var p in engParams) cmdEng.Parameters.Add(p);
                        using (OracleDataAdapter da = new OracleDataAdapter(cmdEng))
                        {
                            da.Fill(dtEng);
                        }
                    }

                    Dictionary<string, List<Tuple<DateTime, DateTime?>>> engDict = new Dictionary<string, List<Tuple<DateTime, DateTime?>>>();
                    foreach (DataRow dr in dtEng.Rows)
                    {
                        string empId = dr["EmpID"].ToString();
                        DateTime sDate = Convert.ToDateTime(dr["StartDate"]);
                        DateTime? eeEndDate = dr["EndDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(dr["EndDate"]) : null;
                        DateTime? cpEndDate = dr["CpEndDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(dr["CpEndDate"]) : null;

                        DateTime? effEnd = null;
                        if (eeEndDate.HasValue && cpEndDate.HasValue)
                            effEnd = eeEndDate.Value < cpEndDate.Value ? eeEndDate.Value : cpEndDate.Value;
                        else if (eeEndDate.HasValue)
                            effEnd = eeEndDate;
                        else if (cpEndDate.HasValue)
                            effEnd = cpEndDate;

                        if (!engDict.ContainsKey(empId))
                            engDict[empId] = new List<Tuple<DateTime, DateTime?>>();
                        engDict[empId].Add(Tuple.Create(sDate, effEnd));
                    }

                    // 3. Fetch Raw Attendance for Year and Month
                    string attSql = @"
                        SELECT EmpID, Day, StatusValue, IsHoliday, LeaveType, AutoSat
                        FROM Attendance
                        WHERE Year = :Year AND Month = :Month
                          AND EmpID != 'GLOBAL' AND EmpID NOT LIKE 'GLOBAL_%'
                          AND Day > 0";
                    DataTable dtAtt = new DataTable();
                    using (OracleCommand cmdAtt = new OracleCommand(attSql, conn))
                    {
                        cmdAtt.BindByName = true;
                        cmdAtt.Parameters.Add(new OracleParameter("Year", year));
                        cmdAtt.Parameters.Add(new OracleParameter("Month", month));
                        using (OracleDataAdapter da = new OracleDataAdapter(cmdAtt))
                        {
                            da.Fill(dtAtt);
                        }
                    }

                    Dictionary<string, Dictionary<int, Tuple<float?, bool, string, int>>> attDict = 
                        new Dictionary<string, Dictionary<int, Tuple<float?, bool, string, int>>>();
                    foreach (DataRow dr in dtAtt.Rows)
                    {
                        string empId = dr["EmpID"].ToString();
                        int day = Convert.ToInt32(dr["Day"]);
                        float? val = dr["StatusValue"] != DBNull.Value ? (float?)Convert.ToSingle(dr["StatusValue"]) : null;
                        bool isHol = dr["IsHoliday"] != DBNull.Value && Convert.ToInt32(dr["IsHoliday"]) == 1;
                        string lType = dr["LeaveType"] != DBNull.Value ? dr["LeaveType"].ToString() : "";
                        int aSat = dr.Table.Columns.Contains("AutoSat") && dr["AutoSat"] != DBNull.Value ? Convert.ToInt32(dr["AutoSat"]) : 0;

                        if (!attDict.ContainsKey(empId))
                            attDict[empId] = new Dictionary<int, Tuple<float?, bool, string, int>>();
                        attDict[empId][day] = Tuple.Create(val, isHol, lType, aSat);
                    }

                    // 4. Iterate and evaluate completeness
                    foreach (DataRow row in dtEmp.Rows)
                    {
                        string masterId = row["MasterId"].ToString();
                        string empName = row["Name"] != DBNull.Value ? row["Name"].ToString() : masterId;
                        DateTime? joinDate = row["JoinDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(row["JoinDate"]) : null;
                        DateTime? resignDate = row["ResignDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(row["ResignDate"]) : null;
                        DateTime? contractEndDate = row["ContractEndDate"] != DBNull.Value ? (DateTime?)Convert.ToDateTime(row["ContractEndDate"]) : null;

                        if (joinDate.HasValue && (year < joinDate.Value.Year || (year == joinDate.Value.Year && month < (joinDate.Value.Month - 1))))
                            continue;
                        if (resignDate.HasValue && (year > resignDate.Value.Year || (year == resignDate.Value.Year && month > (resignDate.Value.Month - 1))))
                            continue;
                        if (contractEndDate.HasValue && (year > contractEndDate.Value.Year || (year == contractEndDate.Value.Year && month > (contractEndDate.Value.Month - 1))))
                            continue;

                        List<Tuple<DateTime, DateTime?>> empEngs = engDict.ContainsKey(masterId) ? engDict[masterId] : new List<Tuple<DateTime, DateTime?>>();
                        Dictionary<int, Tuple<float?, bool, string, int>> empAttMap = attDict.ContainsKey(masterId) ? attDict[masterId] : new Dictionary<int, Tuple<float?, bool, string, int>>();

                        for (int d = 1; d <= daysInMonth; d++)
                        {
                            DateTime currDate = new DateTime(year, dbMonth, d);
                            bool isOutOfBounds = true;
                            foreach (var eng in empEngs)
                            {
                                if (currDate >= eng.Item1.Date && (!eng.Item2.HasValue || currDate <= eng.Item2.Value.Date))
                                {
                                    isOutOfBounds = false;
                                    break;
                                }
                            }

                            if (isOutOfBounds && currDate.DayOfWeek == DayOfWeek.Saturday)
                            {
                                DateTime prevFriday = currDate.AddDays(-1);
                                if (resignDate.HasValue && resignDate.Value.Date == prevFriday.Date && resignDate.Value.DayOfWeek == DayOfWeek.Friday)
                                {
                                    foreach (var eng in empEngs)
                                    {
                                        if (prevFriday >= eng.Item1.Date && (!eng.Item2.HasValue || prevFriday <= eng.Item2.Value.Date))
                                        {
                                            isOutOfBounds = false;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (isOutOfBounds) continue;

                            Tuple<float?, bool, string, int> cell = empAttMap.ContainsKey(d) ? empAttMap[d] : Tuple.Create((float?)null, false, "", 0);

                            // Skip non-holiday Sundays
                            if (currDate.DayOfWeek == DayOfWeek.Sunday && !cell.Item2) continue;

                            string exLabel = $"{empName} ({currDate.ToString("dd-MMM")})";

                            // Condition 1: Missing attendance (non-holiday working day with no record or null value)
                            if (!cell.Item2 && (!empAttMap.ContainsKey(d) || !cell.Item1.HasValue))
                            {
                                result.MissingCount++;
                                if (result.MissingExamples.Count < 3) result.MissingExamples.Add(exLabel);
                            }
                            // Condition 3: Half-day pending pairing
                            else if (cell.Item3.Equals("Pending Pairing", StringComparison.OrdinalIgnoreCase) || 
                                     (cell.Item1.HasValue && cell.Item1.Value == 0.5f && 
                                      !cell.Item3.Equals("Carried", StringComparison.OrdinalIgnoreCase) && 
                                      !cell.Item3.Equals("Paired Paid", StringComparison.OrdinalIgnoreCase) && 
                                      !cell.Item3.Equals("Paired Unpaid", StringComparison.OrdinalIgnoreCase)))
                            {
                                result.PendingPairingCount++;
                                if (result.PendingPairingExamples.Count < 3) result.PendingPairingExamples.Add(exLabel);
                            }
                            // Condition 2: 0 is entered and paid or unpaid is not specified
                            else if (!cell.Item2 && cell.Item1.HasValue && cell.Item1.Value == 0f)
                            {
                                bool isPaidOrUnpaid = cell.Item3.Equals("Paid", StringComparison.OrdinalIgnoreCase) ||
                                                      cell.Item3.Equals("Unpaid", StringComparison.OrdinalIgnoreCase) ||
                                                      cell.Item3.Equals("Paired Unpaid", StringComparison.OrdinalIgnoreCase);

                                if (!isPaidOrUnpaid)
                                {
                                    // Ignore automatic Saturday deduction (AutoSat = 1)
                                    if (currDate.DayOfWeek != DayOfWeek.Saturday || cell.Item4 != 1)
                                    {
                                        result.UnspecifiedZeroCount++;
                                        if (result.UnspecifiedZeroExamples.Count < 3) result.UnspecifiedZeroExamples.Add(exLabel);
                                    }
                                }
                            }
                        }
                    }

                    if (result.MissingCount > 0 || result.UnspecifiedZeroCount > 0 || result.PendingPairingCount > 0)
                    {
                        result.HasIncomplete = true;
                        List<string> msgParts = new List<string>();
                        if (result.MissingCount > 0)
                            msgParts.Add($"{result.MissingCount} day(s) have unentered attendance");
                        if (result.UnspecifiedZeroCount > 0)
                            msgParts.Add($"{result.UnspecifiedZeroCount} absent (0) day(s) have neither Paid nor Unpaid specified");
                        if (result.PendingPairingCount > 0)
                            msgParts.Add($"{result.PendingPairingCount} half-day pairing(s) are pending classification");

                        result.WarningMessage = string.Join("; ", msgParts);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in CheckMonthAttendanceCompleteness: " + ex.Message);
            }

            return result;
        }
    }

    [Serializable]
    public class PocViewRestrictionInfo
    {
        public bool IsRestricted { get; set; }
        public int ViewMode { get; set; }
        public int ViewMonthsAllowed { get; set; }
        public DateTime? ViewCutoffDate { get; set; }
        public int ViewAppliesTo { get; set; }
        public DateTime? MinAllowedDate { get; set; }
        public DateTime? MaxAllowedDate { get; set; }
        public string Description { get; set; }

        public bool IsMonthAllowed(int year, int month1Based)
        {
            if (!IsRestricted) return true;
            DateTime monthStart = new DateTime(year, month1Based, 1);
            DateTime monthEnd = new DateTime(year, month1Based, DateTime.DaysInMonth(year, month1Based));

            if (MinAllowedDate.HasValue && monthEnd < MinAllowedDate.Value)
                return false;
            if (MaxAllowedDate.HasValue && monthStart > MaxAllowedDate.Value)
                return false;

            return true;
        }
    }

    [Serializable]
    public class UserRoleOption
    {
        public string RoleMode { get; set; }     // PrimaryAdmin, SecondaryAdmin, RegularUser, SuperAdmin
        public string Title { get; set; }        // e.g. "Primary Category Admin"
        public string Subtitle { get; set; }     // e.g. "Category Owner for Project"
        public int EffectiveRole { get; set; }   // 1, 0, or 4
        public string Icon { get; set; }         // e.g. "fas fa-user-shield"
        public string BadgeColor { get; set; }   // e.g. "#4f46e5"
    }

    [Serializable]
    public class AttendanceCompletenessResult
    {
        public bool HasIncomplete { get; set; }
        public int MissingCount { get; set; }
        public int UnspecifiedZeroCount { get; set; }
        public int PendingPairingCount { get; set; }
        public List<string> MissingExamples { get; set; }
        public List<string> UnspecifiedZeroExamples { get; set; }
        public List<string> PendingPairingExamples { get; set; }
        public string WarningMessage { get; set; }
    }
}

