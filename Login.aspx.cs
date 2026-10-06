using System;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using Oracle.ManagedDataAccess.Client;

namespace ExcelProcessor
{
    public partial class Login : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                if (User.Identity.IsAuthenticated)
                {
                    Response.Redirect("default.aspx", true);
                    return;
                }

                // Show test bypass indicator if enabled in Web.config
                pnlBypassNotice.Visible = ADHelper.IsBypassEnabled;
            }
        }

        protected void btnLogin_Click(object sender, EventArgs e)
        {
            pnlError.Visible = false;
            lblError.Text = string.Empty;

            string username = (txtUsername.Text ?? "").Trim();
            string password = (txtPassword.Text ?? "").Trim();

            if (string.IsNullOrEmpty(username))
            {
                ShowError("Please enter your domain username or PC number.");
                return;
            }

            if (!ADHelper.IsBypassEnabled && string.IsNullOrEmpty(password))
            {
                ShowError("Please enter your Active Directory password.");
                return;
            }

            string pcno = null;

            // Step 1: Active Directory LDAP Authentication / Bypass
            try
            {
                pcno = ADHelper.AuthenticateAndGetPCNO(username, password);
            }
            catch (Exception ex)
            {
                ShowError("Authentication Failed: " + ex.Message);
                return;
            }

            if (string.IsNullOrWhiteSpace(pcno))
            {
                ShowError("Invalid credentials or employee account not found in Active Directory.");
                return;
            }

            // Step 2: Database Authorization Lookup (Validate that PCNO is in authorized users table)
            bool isAuthorized = false;
            string displayName = pcno;
            string divName = null;

            string connStr = ConfigurationManager.ConnectionStrings["OracleConn"] != null 
                ? ConfigurationManager.ConnectionStrings["OracleConn"].ConnectionString 
                : null;

            if (string.IsNullOrEmpty(connStr))
            {
                ShowError("System Configuration Error: Database connection string 'OracleConn' is missing.");
                return;
            }

            string targetTable = ConfigurationManager.AppSettings["AuthUsersTable"] ?? "Excel_AppUsers";

            try
            {
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    conn.Open();

                    // Primary check against configured table (e.g. Excel_AppUsers)
                    bool checkedPrimary = false;
                    try
                    {
                        string sql = "SELECT PCNO, NAME, DIVNAME FROM " + targetTable + " WHERE UPPER(TRIM(PCNO)) = UPPER(TRIM(:PCNO)) AND ROWNUM <= 1";
                        using (OracleCommand cmd = new OracleCommand(sql, conn))
                        {
                            cmd.Parameters.Add(new OracleParameter("PCNO", pcno));
                            using (OracleDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    isAuthorized = true;
                                    if (!reader.IsDBNull(1))
                                    {
                                        displayName = reader.GetString(1);
                                    }
                                    if (!reader.IsDBNull(2))
                                    {
                                        divName = reader.GetString(2);
                                    }
                                }
                                checkedPrimary = true;
                            }
                        }
                    }
                    catch
                    {
                        // Retry with PCNO, NAME if DIVNAME column isn't in target table
                        try
                        {
                            string sqlBasic = "SELECT PCNO, NAME FROM " + targetTable + " WHERE UPPER(TRIM(PCNO)) = UPPER(TRIM(:PCNO)) AND ROWNUM <= 1";
                            using (OracleCommand cmdBasic = new OracleCommand(sqlBasic, conn))
                            {
                                cmdBasic.Parameters.Add(new OracleParameter("PCNO", pcno));
                                using (OracleDataReader readerBasic = cmdBasic.ExecuteReader())
                                {
                                    if (readerBasic.Read())
                                    {
                                        isAuthorized = true;
                                        if (!readerBasic.IsDBNull(1))
                                        {
                                            displayName = readerBasic.GetString(1);
                                        }
                                    }
                                    checkedPrimary = true;
                                }
                            }
                        }
                        catch
                        {
                            checkedPrimary = false;
                        }
                    }

                    // Fallback check against APP_USERS table if primary table query had column differences
                    if (!checkedPrimary)
                    {
                        try
                        {
                            string sqlAlt = "SELECT PCNO, DISPLAY_NAME FROM APP_USERS WHERE UPPER(TRIM(PCNO)) = UPPER(TRIM(:PCNO)) AND ROWNUM <= 1";
                            using (OracleCommand cmdAlt = new OracleCommand(sqlAlt, conn))
                            {
                                cmdAlt.Parameters.Add(new OracleParameter("PCNO", pcno));
                                using (OracleDataReader readerAlt = cmdAlt.ExecuteReader())
                                {
                                    if (readerAlt.Read())
                                    {
                                        isAuthorized = true;
                                        if (!readerAlt.IsDBNull(1))
                                        {
                                            displayName = readerAlt.GetString(1);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception exAlt)
                        {
                            throw new Exception("Unable to verify user authorization in database: " + exAlt.Message, exAlt);
                        }
                    }

                    // If DivName is still missing, lookup in hrdata.empdetails or empdetails
                    if (isAuthorized && string.IsNullOrWhiteSpace(divName))
                    {
                        try
                        {
                            string sqlEmp = "SELECT DIVNAME, NAME FROM hrdata.empdetails WHERE UPPER(TRIM(PCNO)) = UPPER(TRIM(:PCNO)) AND ROWNUM <= 1";
                            using (OracleCommand cmdEmp = new OracleCommand(sqlEmp, conn))
                            {
                                cmdEmp.Parameters.Add(new OracleParameter("PCNO", pcno));
                                using (OracleDataReader readerEmp = cmdEmp.ExecuteReader())
                                {
                                    if (readerEmp.Read())
                                    {
                                        if (!readerEmp.IsDBNull(0))
                                            divName = readerEmp.GetString(0);
                                        if (string.IsNullOrWhiteSpace(displayName) && !readerEmp.IsDBNull(1))
                                            displayName = readerEmp.GetString(1);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            try
                            {
                                string sqlEmp2 = "SELECT DIVNAME, NAME FROM empdetails WHERE UPPER(TRIM(PCNO)) = UPPER(TRIM(:PCNO)) AND ROWNUM <= 1";
                                using (OracleCommand cmdEmp2 = new OracleCommand(sqlEmp2, conn))
                                {
                                    cmdEmp2.Parameters.Add(new OracleParameter("PCNO", pcno));
                                    using (OracleDataReader readerEmp2 = cmdEmp2.ExecuteReader())
                                    {
                                        if (readerEmp2.Read())
                                        {
                                            if (!readerEmp2.IsDBNull(0))
                                                divName = readerEmp2.GetString(0);
                                            if (string.IsNullOrWhiteSpace(displayName) && !readerEmp2.IsDBNull(1))
                                                displayName = readerEmp2.GetString(1);
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception exDb)
            {
                ShowError("Database Verification Error: " + exDb.Message);
                return;
            }

            // Step 3: Grant or Deny Access
            if (!isAuthorized)
            {
                ShowError("Access Denied: PC number '" + pcno + "' is not authorized to access this system. Please contact the administrator.");
                return;
            }

            // Grant session access and authentication cookie
            Session["UserPCNO"] = pcno;
            Session["UserName"] = string.IsNullOrWhiteSpace(displayName) ? pcno : displayName;
            Session["UserDivName"] = string.IsNullOrWhiteSpace(divName) ? pcno : divName;

            FormsAuthentication.SetAuthCookie(pcno, false);
            Response.Redirect("default.aspx", true);
        }

        private void ShowError(string message)
        {
            lblError.Text = message;
            pnlError.Visible = true;
        }
    }
}
