using System;
using System.Data;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.UI.WebControls;
using AttendanceApp.Utils;
using Oracle.ManagedDataAccess.Client;

namespace AttendanceApp
{
    public partial class Login : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                if (User.Identity.IsAuthenticated)
                {
                    Response.Redirect("Dashboard.aspx");
                }
            }
        }

        protected void btnLogin_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();

            if (string.IsNullOrEmpty(username))
            {
                ShowError("Username is required.");
                return;
            }

            string pcno = null;

            try
            {
                pcno = ADHelper.AuthenticateAndGetPCNO(username, password);
            }
            catch (Exception ex)
            {
                ShowError("AD Error: " + ex.Message);
                return;
            }

            if (string.IsNullOrEmpty(pcno))
            {
                ShowError("Invalid credentials or user not found in AD.");
                return;
            }

            // Detect all available access roles for this user
            List<UserRoleOption> roles = null;
            try
            {
                roles = DBHelper.GetAvailableUserRoles(pcno);
            }
            catch (Exception ex)
            {
                ShowError("Database Connection Error: Could not connect to Attendance Database. Please verify database services are running. " + ex.Message);
                return;
            }

            if (roles == null || roles.Count == 0)
            {
                // Inspect AppUsers to provide accurate feedback
                try
                {
                    string queryRoles = "SELECT DISTINCT Role FROM AppUsers WHERE PCNO = :PCNO";
                    DataTable dtRoles = DBHelper.ExecuteQuery(DBHelper.GetAttendanceDBConnection(), queryRoles, new OracleParameter("PCNO", pcno));
                    if (dtRoles == null || dtRoles.Rows.Count == 0)
                    {
                        ShowError("Access denied. You are not authorized to log in. Please contact an administrator.");
                        return;
                    }

                    List<int> allRoles = new List<int>();
                    foreach (DataRow r in dtRoles.Rows)
                    {
                        if (r["Role"] != DBNull.Value) allRoles.Add(Convert.ToInt32(r["Role"]));
                    }

                    if (allRoles.Contains(5) && allRoles.Count == 1)
                    {
                        ShowError("Access denied. Your super administrator access has been revoked. Please contact a system administrator.");
                    }
                    else if (allRoles.Contains(2) && allRoles.Count == 1)
                    {
                        ShowError("Access denied. Your administrator access has been revoked. Please contact a system administrator.");
                    }
                    else if (allRoles.Contains(3) && allRoles.Count == 1)
                    {
                        ShowError("Access denied. Your regular user (POC) access has been revoked. Please contact a system administrator.");
                    }
                    else if (allRoles.Contains(7) && allRoles.Count == 1)
                    {
                        ShowError("Access denied. Your sub user access has been revoked. Please contact a system administrator.");
                    }
                    else
                    {
                        ShowError("Access denied. Your access has been revoked or no active roles are assigned to your account. Please contact a system administrator.");
                    }
                }
                catch
                {
                    ShowError("Access denied. No active roles found for your account.");
                }
                return;
            }

            // Store temporary login context in session
            Session["PendingPCNO"] = pcno;
            Session["PendingRoles"] = roles;

            if (roles.Count == 1)
            {
                // Single role available: complete login directly
                CompleteUserLogin(pcno, roles[0], roles);
            }
            else
            {
                // Multiple roles available: show interactive Role Selection Panel
                string displayName = pcno;
                try
                {
                    string nameQuery = "SELECT Name FROM AppUsers WHERE PCNO = :PCNO AND ROWNUM <= 1";
                    object nameRes = DBHelper.ExecuteScalar(DBHelper.GetAttendanceDBConnection(), nameQuery, new OracleParameter("PCNO", pcno));
                    if (nameRes != null && nameRes != DBNull.Value && !string.IsNullOrEmpty(nameRes.ToString()))
                    {
                        displayName = nameRes.ToString();
                    }
                }
                catch { }

                lblRoleSelectionUserName.Text = displayName;
                rptRoleOptions.DataSource = roles;
                rptRoleOptions.DataBind();

                pnlLoginForm.Visible = false;
                pnlRoleSelection.Visible = true;
                lblError.Visible = false;
            }
        }

        protected void rptRoleOptions_ItemCommand(object source, RepeaterCommandEventArgs e)
        {
            if (e.CommandName == "SelectRole")
            {
                string selectedMode = e.CommandArgument != null ? e.CommandArgument.ToString() : "";
                string pcno = Session["PendingPCNO"] as string;
                List<UserRoleOption> roles = Session["PendingRoles"] as List<UserRoleOption>;

                if (string.IsNullOrEmpty(pcno) || roles == null)
                {
                    ShowError("Session expired. Please log in again.");
                    pnlRoleSelection.Visible = false;
                    pnlLoginForm.Visible = true;
                    return;
                }

                UserRoleOption targetRole = roles.Find(r => r.RoleMode == selectedMode);
                if (targetRole == null)
                {
                    targetRole = roles[0];
                }

                CompleteUserLogin(pcno, targetRole, roles);
            }
        }

        protected void btnCancelRoleSelection_Click(object sender, EventArgs e)
        {
            Session["PendingPCNO"] = null;
            Session["PendingRoles"] = null;
            pnlRoleSelection.Visible = false;
            pnlLoginForm.Visible = true;
            lblError.Visible = false;
        }

        private void CompleteUserLogin(string pcno, UserRoleOption selectedRole, List<UserRoleOption> availableRoles)
        {
            // Load allowed divisions for regular user and sub user
            List<string> allowedDivisions = new List<string>();
            if (selectedRole.EffectiveRole != 1 && selectedRole.EffectiveRole != 4)
            {
                try
                {
                    DBHelper.EnsureDivisionsTableExists();
                    string queryDivs = "SELECT DivisionName FROM UserDivisions WHERE PCNO = :PCNO";
                    DataTable dtDivs = DBHelper.ExecuteQuery(DBHelper.GetAttendanceDBConnection(), queryDivs, new OracleParameter("PCNO", pcno));
                    foreach (DataRow row in dtDivs.Rows)
                    {
                        allowedDivisions.Add(row["DivisionName"].ToString());
                    }
                    if (allowedDivisions.Count == 0)
                    {
                        allowedDivisions.Add("D-USER");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error fetching user divisions: " + ex.Message);
                }
            }

            // Fetch Name/Division from CompanyDB (hrdata.empdetails)
            string division = "";
            string name = "User";
            string designation = "";
            try
            {
                string queryDiv = "SELECT NAME, DESIGNATION, DIVNAME FROM hrdata.empdetails WHERE PCNO = :PCNO AND ROWNUM <= 1";
                DataTable dt = DBHelper.ExecuteQuery(DBHelper.GetCompanyDBConnection(), queryDiv, new OracleParameter("PCNO", pcno));
                if (dt.Rows.Count > 0)
                {
                    name        = dt.Rows[0]["NAME"].ToString();
                    designation = dt.Rows[0]["DESIGNATION"].ToString();
                    division    = dt.Rows[0]["DIVNAME"].ToString();
                }
                else
                {
                    if (selectedRole.EffectiveRole == 1 || selectedRole.EffectiveRole == 4)
                    {
                        name        = selectedRole.EffectiveRole == 4 ? "Super Admin" : "Admin";
                        designation = selectedRole.EffectiveRole == 4 ? "Super Administrator" : "System Administrator";
                        division    = "D-ADMIN";
                    }
                    else
                    {
                        try
                        {
                            string queryName = "SELECT Name FROM AppUsers WHERE PCNO = :PCNO AND ROWNUM <= 1";
                            object nameRes = DBHelper.ExecuteScalar(DBHelper.GetAttendanceDBConnection(), queryName, new OracleParameter("PCNO", pcno));
                            name = (nameRes != null && nameRes != DBNull.Value) ? nameRes.ToString() : pcno;
                        }
                        catch { name = pcno; }
                        designation = "";
                        division    = allowedDivisions.Count > 0 ? allowedDivisions[0] : "";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error connecting to HR DB: " + ex.Message);
            }

            // Store in Session
            Session["PCNO"] = pcno;
            Session["Role"] = selectedRole.EffectiveRole;
            Session["RoleMode"] = selectedRole.RoleMode;
            Session["UserRoles"] = availableRoles;
            Session["Name"] = name;
            Session["Designation"] = designation;
            Session["AllowedDivisions"] = allowedDivisions;

            if (selectedRole.EffectiveRole == 1 || selectedRole.EffectiveRole == 4)
            {
                string divPrefix = string.IsNullOrEmpty(division) ? "D-ADMIN" : division;
                if (divPrefix.Contains("/"))
                {
                    divPrefix = divPrefix.Split('/')[0].Trim();
                }
                Session["Division"] = divPrefix;
            }
            else
            {
                Session["Division"] = allowedDivisions.Count > 0 ? allowedDivisions[0] : "D-USER";
            }

            try
            {
                string updateNameQuery = "UPDATE AppUsers SET Name = :Name WHERE PCNO = :PCNO";
                int affected = DBHelper.ExecuteNonQuery(DBHelper.GetAttendanceDBConnection(), updateNameQuery,
                    new OracleParameter("Name", name),
                    new OracleParameter("PCNO", pcno));
                if (affected == 0)
                {
                    string insertQuery = "INSERT INTO AppUsers (PCNO, Name, Role) VALUES (:PCNO, :Name, :Role)";
                    DBHelper.ExecuteNonQuery(DBHelper.GetAttendanceDBConnection(), insertQuery,
                        new OracleParameter("PCNO", pcno),
                        new OracleParameter("Name", name),
                        new OracleParameter("Role", selectedRole.EffectiveRole));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error updating user name on login: " + ex.Message);
            }

            FormsAuthentication.SetAuthCookie(pcno, false);
            Response.Redirect("Dashboard.aspx");
        }

        private void ShowError(string message)
        {
            lblError.Text = message;
            lblError.Visible = true;
        }
    }
}
