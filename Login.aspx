<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="ExcelProcessor.Login" %>
<!DOCTYPE html>
<html lang="en">
<head runat="server">
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
  <meta name="description" content="Secure Enterprise Login - Employee Record Mapper" />
  <title>Sign In - Employee Record Mapper</title>
  
  <!-- Local Offline Stylesheets -->
  <link rel="stylesheet" href="lib/bootstrap.min.css" />
  <link rel="stylesheet" href="style.css" />

  <style>
    body {
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
      margin: 0;
      padding: 24px;
      background-color: var(--bg-color, #f8fafc);
      background-image: 
        radial-gradient(at 0% 0%, rgba(37, 99, 235, 0.05) 0px, transparent 50%),
        radial-gradient(at 100% 0%, rgba(14, 165, 233, 0.06) 0px, transparent 50%),
        radial-gradient(at 50% 100%, rgba(99, 102, 241, 0.04) 0px, transparent 60%);
      font-family: var(--font-stack, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif);
      color: #0f172a;
    }

    .login-wrapper {
      width: 100%;
      max-width: 440px;
    }

    .login-card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 16px;
      padding: 36px 32px;
      box-shadow: 0 4px 6px -1px rgba(15, 23, 42, 0.05), 0 20px 25px -5px rgba(15, 23, 42, 0.06);
      transition: all 0.25s ease;
    }

    .login-brand {
      display: flex;
      flex-direction: column;
      align-items: center;
      text-align: center;
      margin-bottom: 28px;
    }

    .login-icon-box {
      width: 54px;
      height: 54px;
      border-radius: 14px;
      background: #eff6ff;
      border: 1px solid #bfdbfe;
      display: flex;
      align-items: center;
      justify-content: center;
      margin-bottom: 16px;
      box-shadow: 0 4px 12px rgba(37, 99, 235, 0.1);
    }

    .login-title {
      font-size: 1.35rem;
      font-weight: 700;
      color: #0f172a;
      margin: 0 0 6px 0;
      letter-spacing: -0.01em;
    }

    .login-subtitle {
      font-size: 0.86rem;
      color: #64748b;
      margin: 0;
      line-height: 1.45;
    }

    .form-group-custom {
      margin-bottom: 20px;
    }

    .form-label-custom {
      display: block;
      font-size: 0.82rem;
      font-weight: 600;
      color: #334155;
      margin-bottom: 7px;
    }

    .input-icon-wrapper {
      position: relative;
    }

    .input-icon {
      position: absolute;
      left: 14px;
      top: 50%;
      transform: translateY(-50%);
      color: #94a3b8;
      display: flex;
      align-items: center;
      pointer-events: none;
    }

    .form-control-custom {
      width: 100%;
      height: 44px;
      padding: 10px 14px 10px 42px;
      font-size: 0.9rem;
      border: 1.5px solid #cbd5e1;
      border-radius: 10px;
      background-color: #f8fafc;
      color: #0f172a;
      transition: all 0.2s ease;
      box-sizing: border-box;
    }

    .form-control-custom:focus {
      outline: none;
      border-color: #2563eb;
      background-color: #ffffff;
      box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
    }

    .btn-submit {
      width: 100%;
      height: 44px;
      background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
      color: #ffffff;
      border: none;
      border-radius: 10px;
      font-size: 0.92rem;
      font-weight: 600;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: all 0.2s ease;
      box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
      margin-top: 24px;
    }

    .btn-submit:hover {
      background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%);
      transform: translateY(-1px);
      box-shadow: 0 6px 16px rgba(37, 99, 235, 0.32);
    }

    .btn-submit:active {
      transform: translateY(0);
    }

    .alert-custom-error {
      background: #fff1f2;
      border: 1px solid #fecdd3;
      color: #be123c;
      padding: 12px 14px;
      border-radius: 10px;
      font-size: 0.84rem;
      font-weight: 500;
      margin-bottom: 20px;
      display: flex;
      align-items: flex-start;
      gap: 10px;
      line-height: 1.4;
    }

    .badge-bypass-mode {
      background: #fef3c7;
      border: 1px solid #fde68a;
      color: #92400e;
      padding: 10px 14px;
      border-radius: 10px;
      font-size: 0.82rem;
      margin-bottom: 22px;
      display: flex;
      align-items: center;
      gap: 8px;
      line-height: 1.4;
    }

    .footer-note {
      text-align: center;
      margin-top: 24px;
      padding-top: 20px;
      border-top: 1px solid #f1f5f9;
      font-size: 0.77rem;
      color: #94a3b8;
    }
  </style>
</head>
<body>
  <div class="login-wrapper">
    <div class="login-card">
      
      <!-- Brand & Title -->
      <div class="login-brand">
        <div class="login-icon-box">
          <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
            <line x1="16" y1="13" x2="8" y2="13"></line>
            <line x1="16" y1="17" x2="8" y2="17"></line>
            <polyline points="10 9 9 9 8 9"></polyline>
          </svg>
        </div>
        <h1 class="login-title">Employee Record Mapper</h1>
        <p class="login-subtitle">Sign in with your Domain / Active Directory account</p>
      </div>

      <form id="loginForm" runat="server">

        <!-- Bypass Test Mode Notice -->
        <asp:Panel ID="pnlBypassNotice" runat="server" Visible="false" CssClass="badge-bypass-mode">
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" style="flex-shrink: 0;">
            <path d="M8.982 1.566a1.13 1.13 0 0 0-1.96 0L.165 13.233c-.457.778.091 1.767.98 1.767h13.713c.889 0 1.438-.99.98-1.767L8.982 1.566zM8 5c.535 0 .954.462.9.995l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 5.995A.905.905 0 0 1 8 5zm.002 6a1 1 0 1 1 0 2 1 1 0 0 1 0-2z"/>
          </svg>
          <div>
            <strong>Test Bypass Active</strong>: LDAP bypass is enabled in Web.config. You can sign in using test PC numbers (e.g., <code>1001</code> or <code>1002</code>).
          </div>
        </asp:Panel>

        <!-- Error Notification Panel -->
        <asp:Panel ID="pnlError" runat="server" Visible="false" CssClass="alert-custom-error">
          <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="currentColor" viewBox="0 0 16 16" style="flex-shrink: 0; margin-top: 1px;">
            <path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z"/>
            <path d="M7.002 11a1 1 0 1 1 2 0 1 1 0 0 1-2 0zM7.1 4.995a.905.905 0 1 1 1.8 0l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 4.995z"/>
          </svg>
          <asp:Label ID="lblError" runat="server"></asp:Label>
        </asp:Panel>

        <!-- Username Input -->
        <div class="form-group-custom">
          <label for="txtUsername" class="form-label-custom">Domain Username / PC Number</label>
          <div class="input-icon-wrapper">
            <span class="input-icon">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                <circle cx="12" cy="7" r="4"></circle>
              </svg>
            </span>
            <asp:TextBox ID="txtUsername" runat="server" CssClass="form-control-custom" placeholder="e.g. jdoe or 1001" autocomplete="username"></asp:TextBox>
          </div>
        </div>

        <!-- Password Input -->
        <div class="form-group-custom">
          <label for="txtPassword" class="form-label-custom">Password</label>
          <div class="input-icon-wrapper">
            <span class="input-icon">
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24">
                <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
              </svg>
            </span>
            <asp:TextBox ID="txtPassword" runat="server" CssClass="form-control-custom" TextMode="Password" placeholder="••••••••••••" autocomplete="current-password"></asp:TextBox>
          </div>
        </div>

        <!-- Submit Button -->
        <asp:Button ID="btnLogin" runat="server" Text="Sign In to System" CssClass="btn-submit" OnClick="btnLogin_Click" />

      </form>

      <!-- Footer Security Note -->
      <div class="footer-note">
        Authorized Access Only &bull; Access is audited &bull; DRDO LRDE
      </div>

    </div>
  </div>

  <script>
    window.addEventListener('DOMContentLoaded', function() {
      var u = document.getElementById('<%= txtUsername.ClientID %>');
      var p = document.getElementById('<%= txtPassword.ClientID %>');
      if (u) {
        if (!u.value) {
          u.focus();
        } else if (p) {
          p.focus();
        }
      }
    });
  </script>
</body>
</html>
