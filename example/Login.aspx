<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="AttendanceApp.Login" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
    <meta name="description" content="" />
    <meta name="author" content="" />

    <title>ARS - Attendance Recording System (Skilled, Semi-Skilled) Login</title>
    <!-- Custom fonts for this template-->
    <link href="Static/fontawesome-free/css/all.min.css" rel="stylesheet" type="text/css" />
    <!-- Custom styles for this template-->
    <link href="Static/css/sb-admin-2.min.css" rel="stylesheet" />
    <style>
        .role-selection-tile {
            display: flex;
            align-items: center;
            justify-content: space-between;
            width: 100%;
            background: #ffffff;
            border: 1.5px solid #e2e8f0;
            border-radius: 12px;
            padding: 12px 14px;
            margin-bottom: 10px;
            text-decoration: none !important;
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
            cursor: pointer;
            outline: none;
        }
        .role-selection-tile:hover {
            border-color: #4e73df;
            background-color: #f8faff;
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(78, 115, 223, 0.12);
        }
        .role-selection-tile:active {
            transform: translateY(0);
        }
        .role-tile-left {
            display: flex;
            align-items: center;
            min-width: 0;
            flex: 1 1 auto;
            margin-right: 10px;
        }
        .role-tile-avatar {
            width: 40px;
            height: 40px;
            min-width: 40px;
            border-radius: 10px;
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1rem;
            margin-right: 12px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.08);
        }
        .role-tile-info {
            min-width: 0;
            text-align: left;
        }
        .role-tile-title {
            font-weight: 700;
            font-size: 0.92rem;
            color: #1e293b;
            line-height: 1.3;
            margin-bottom: 2px;
            transition: color 0.15s ease;
        }
        .role-selection-tile:hover .role-tile-title {
            color: #2e59d9;
        }
        .role-tile-subtitle {
            font-size: 0.76rem;
            color: #64748b;
            line-height: 1.3;
        }
        .role-tile-action {
            flex: 0 0 auto;
        }
        .role-tile-arrow {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            background: #f1f5f9;
            color: #64748b;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.75rem;
            transition: all 0.2s ease;
        }
        .role-selection-tile:hover .role-tile-arrow {
            background: #4e73df;
            color: #ffffff;
            transform: translateX(3px);
        }
    </style>
</head>
<body class="bg-gradient-primary">
    <form id="form1" runat="server">
        <div class="container" style="margin-top: 8%">

            <!-- Outer Row -->
            <div class="row justify-content-center">

                <div class="col-xl-10 col-lg-12 col-md-9">

                    <div class="card o-hidden border-0 shadow-lg my-5" style="border-radius: 16px;">
                        <div class="card-body p-0">
                            <!-- Nested Row within Card Body -->
                            <div class="row">
                                <div class="col-lg-6 d-none d-lg-flex align-items-center justify-content-center bg-light" style="border-radius: 16px 0 0 16px; padding: 40px 30px;">
                                    <div class="text-center w-100" style="max-width: 320px;">
                                        <img src="Static/Images/newlrdelogo.png" style="max-width: 240px; width: 100%; height: auto; filter: drop-shadow(0 4px 12px rgba(0,0,0,0.06));" alt="LRDE Logo" />
                                        <div class="mt-4">
                                            <h4 class="font-weight-bold text-gray-900 mb-1" style="letter-spacing: 0.8px; font-size: 1.35rem;">ARS</h4>
                                            <p class="text-muted small mb-0 font-weight-bold" style="font-size: 0.82rem; line-height: 1.35;">Attendance Recording System (Skilled, Semi-Skilled)</p>
                                        </div>
                                    </div>
                                </div>
                                <div class="col-lg-6">
                                    <div class="p-5">
                                        <asp:Label ID="lblError" runat="server" CssClass="text-danger d-block mb-3 font-weight-bold text-center" Visible="false"></asp:Label>

                                        <asp:Panel ID="pnlLoginForm" runat="server">
                                            <div class="text-center">
                                                <h1 class="h4 text-gray-900 font-weight-bold mb-1">Welcome Back</h1>
                                                <p class="text-muted small mb-4">Please log in with your domain credentials</p>
                                            </div>
                                            <div class="user">
                                                <div class="form-group">
                                                    <asp:TextBox ID="txtUsername" runat="server" CssClass="form-control form-control-user" placeholder="Enter your Username..."></asp:TextBox>
                                                </div>
                                                <div class="form-group">
                                                    <asp:TextBox ID="txtPassword" runat="server" CssClass="form-control form-control-user" TextMode="Password" placeholder="Password"></asp:TextBox>
                                                </div>
                                                <asp:Button ID="btnLogin" CssClass="btn btn-primary btn-user btn-block" OnClick="btnLogin_Click" runat="server" Text="Login" />

                                                <hr />

                                                <a href="http://www.lrde.com" class="btn btn-facebook btn-user btn-block">
                                                    <i class="fas fa-home"></i>LRDE Home
                                                </a>
                                            </div>
                                        </asp:Panel>

                                        <asp:Panel ID="pnlRoleSelection" runat="server" Visible="false">
                                            <div class="text-center mb-4">
                                                <div class="d-inline-flex align-items-center justify-content-center px-3 py-1 mb-2" style="background-color: #eef2ff; color: #4338ca; border-radius: 20px; font-size: 0.8rem; font-weight: 600;">
                                                    <i class="fas fa-user-shield mr-1.5" style="margin-right: 6px;"></i> Role Selection
                                                </div>
                                                <h4 class="text-gray-900 font-weight-bold mb-1" style="font-size: 1.25rem;">Select Access Role</h4>
                                                <p class="text-muted small mb-0">Welcome, <asp:Label ID="lblRoleSelectionUserName" runat="server" Font-Bold="true" CssClass="text-gray-900"></asp:Label>! Choose your active role:</p>
                                            </div>

                                            <div class="role-selection-list mb-3">
                                                <asp:Repeater ID="rptRoleOptions" runat="server" OnItemCommand="rptRoleOptions_ItemCommand">
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="btnSelectRoleTile" runat="server" CommandName="SelectRole" CommandArgument='<%# Eval("RoleMode") %>' CssClass="role-selection-tile">
                                                            <div class="role-tile-left">
                                                                <div class="role-tile-avatar" style='<%# "background-color:" + Eval("BadgeColor") %>'>
                                                                    <i class='<%# Eval("Icon") %>'></i>
                                                                </div>
                                                                <div class="role-tile-info">
                                                                    <div class="role-tile-title"><%# Eval("Title") %></div>
                                                                    <div class="role-tile-subtitle"><%# Eval("Subtitle") %></div>
                                                                </div>
                                                            </div>
                                                            <div class="role-tile-action">
                                                                <div class="role-tile-arrow">
                                                                    <i class="fas fa-arrow-right"></i>
                                                                </div>
                                                            </div>
                                                        </asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:Repeater>
                                            </div>

                                            <div class="text-center mt-3">
                                                <asp:LinkButton ID="btnCancelRoleSelection" runat="server" OnClick="btnCancelRoleSelection_Click" CssClass="small text-secondary font-weight-bold" style="text-decoration: none;">
                                                    <i class="fas fa-arrow-left mr-1"></i> Back to Login
                                                </asp:LinkButton>
                                            </div>
                                        </asp:Panel>
                                        <hr />
                                        <div class="text-center">
                                            <p class="small" style="color:black">For any Help/Feedback please mail ITISG@(it-soft@lrde.com)</p>
                                        </div>

                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="text-center"><p class="small" style="color:black">Designed and Developed By D-KRM/ITISG</p></div>
                    </div>
                    
                </div>

            </div>
            

        </div>

        <!-- Bootstrap core JavaScript-->
        <script src="Static/jquery/jquery.min.js"></script>
        <script src="Static/bootstrap/js/bootstrap.bundle.min.js"></script>

        <!-- Core plugin JavaScript-->
        <script src="Static/jquery-easing/jquery.easing.min.js"></script>

        <!-- Custom scripts for all pages-->
        <script src="Static/js/sb-admin-2.min.js"></script>

        <!-- On Page load Username should be focused-->
        <script>
            $(document).ready(function () {
                try {
                    var _t = localStorage.getItem('app-theme');
                    localStorage.clear();
                    if (_t) localStorage.setItem('app-theme', _t);
                    sessionStorage.clear();
                } catch (e) {
                    console.error("Failed to clear local/session storage:", e);
                }

                $('#txtUsername').focus();

                $('#txtUsername').keypress(function (e) {
                    if (e.which == 13) {
                        e.preventDefault();
                        $('#txtPassword').focus();
                    }
                });
            });
        </script>
    </form>
</body>
</html>
