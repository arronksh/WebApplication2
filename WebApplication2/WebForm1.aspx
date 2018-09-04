<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="WebApplication2.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
        </div>
        <asp:FileUpload ID="FileUpload1" runat="server" Width="1103px" /> <asp:Button ID="Button2" runat="server" Text="Cek nama file" /><br />
        <asp:Button ID="Button1" runat="server" Text="Button" />
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:gridview id="GridView1" runat="server" xmlns:asp="#unknown" />
        <br />
        <asp:Button ID="btnLogout" runat="server" Text="Logout" />
    </form>
</body>
</html>
