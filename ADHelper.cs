using System;
using System.Configuration;
using System.DirectoryServices;

namespace ExcelProcessor
{
    public static class ADHelper
    {
        public static bool IsBypassEnabled
        {
            get
            {
                return string.Equals(ConfigurationManager.AppSettings["EnableLdapBypass"], "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        public static string AuthenticateAndGetPCNO(string username, string password)
        {
            if (string.IsNullOrWhiteSpace(username))
                throw new ArgumentException("Username cannot be empty.");

            // Check if LDAP bypass is enabled in Web.config
            if (IsBypassEnabled)
            {
                // Local testing bypass:
                string lower = username.Trim().ToLowerInvariant();
                if (lower == "admin" || lower == "aadmin")
                    return "1001";
                if (lower == "test" || lower == "user")
                    return "1002";
                
                // Allow using any numeric or alphanumeric PCNO directly
                return username.Trim();
            }

            // Production Mode: Active Directory LDAP Authentication
            if (string.IsNullOrWhiteSpace(password))
                throw new ArgumentException("Password cannot be empty.");

            string pcno = null;
            string ldapPath = ConfigurationManager.AppSettings["ADConnectionPath"];

            if (string.IsNullOrWhiteSpace(ldapPath))
                throw new InvalidOperationException("ADConnectionPath is not configured in Web.config.");

            try
            {
                // In C# DirectoryEntry doesn't validate credentials until NativeObject or a search is performed
                using (DirectoryEntry entry = new DirectoryEntry(ldapPath, username.Trim(), password))
                {
                    object native = entry.NativeObject; // Forces credentials authentication against AD

                    using (DirectorySearcher search = new DirectorySearcher(entry))
                    {
                        search.Filter = "(SAMAccountName=" + username.Trim() + ")";
                        search.PropertiesToLoad.Add("EmployeeID");
                        SearchResult result = search.FindOne();

                        if (result != null)
                        {
                            using (DirectoryEntry dsresult = result.GetDirectoryEntry())
                            {
                                if (dsresult.Properties.Contains("EmployeeID") && dsresult.Properties["EmployeeID"].Count > 0)
                                {
                                    pcno = dsresult.Properties["EmployeeID"][0].ToString();
                                }
                                else
                                {
                                    // Fallback: If EmployeeID property is not populated, use username
                                    pcno = username.Trim();

                                    if (lowerFallbackAdmin(username))
                                    {
                                        pcno = "1001";
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw new Exception("Authentication succeeded, but user account was not found in Active Directory search.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Active Directory Authentication Error: " + ex.Message, ex);
            }

            return pcno;
        }

        private static bool lowerFallbackAdmin(string username)
        {
            string lower = username.ToLowerInvariant();
            return lower == "aadmin" || lower == "admin";
        }
    }
}
