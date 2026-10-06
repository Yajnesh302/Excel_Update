using System;
using System.DirectoryServices;
using System.Configuration;

namespace AttendanceApp.Utils
{
    public static class ADHelper
    {
        public static string AuthenticateAndGetPCNO(string username, string password)
        {
            // Local offline testing bypass
            if (username == "1001" || username.ToLower() == "admin" || username.ToLower() == "aadmin")
            {
                return "1001";
            }
            if (username == "1002")
            {
                return "1002";
            }
            if (username == "1003")
            {
                return "1003";
            }
            if (username == "1004")
            {
                return "1004";
            }


            string pcno = null;
            string ldapPath = ConfigurationManager.AppSettings["ADConnectionPath"];
            
            // In C# DirectoryEntry doesn't validate credentials until NativeObject or a search is performed.
            using (DirectoryEntry entry = new DirectoryEntry(ldapPath, username, password))
            {
                object native = entry.NativeObject; // Forces authentication

                using (DirectorySearcher search = new DirectorySearcher(entry))
                {
                    search.Filter = "(SAMAccountName=" + username + ")";
                    search.PropertiesToLoad.Add("EmployeeID");
                    SearchResult result = search.FindOne();
                    
                    if (result != null)
                    {
                        using (DirectoryEntry dsresult = result.GetDirectoryEntry())
                        {
                            if (dsresult.Properties.Contains("EmployeeID"))
                            {
                                pcno = dsresult.Properties["EmployeeID"][0].ToString();
                            }
                            else
                            {
                                // Fallback for local AD testing if EmployeeID is not configured
                                pcno = username; 
                                
                                // If testing aadmin, treat as Admin PCNO 1001
                                if (username.ToLower() == "aadmin" || username.ToLower() == "admin")
                                {
                                    pcno = "1001";
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("Authentication succeeded, but user was not found in AD search.");
                    }
                }
            }
            
            return pcno;
        }
    }
}
