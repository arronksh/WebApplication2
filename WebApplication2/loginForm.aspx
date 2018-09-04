<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="loginForm.aspx.vb" Inherits="WebApplication2.loginForm" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="Label1" runat="server" Text="ID"></asp:Label>
            <asp:textbox runat="server" ID="tbID"></asp:textbox>
            <br />
            <asp:Label ID="Label2" runat="server" Text="PASS"></asp:Label>
            <asp:textbox runat="server" ID="tbPass"></asp:textbox>
            <br />
            <asp:Button ID="BTNLOGIN" runat="server" Text="LOGIN" Width="222px" />
        </div>
</body>
</html>
    </form>

