<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Develpment.aspx.cs" Inherits="Develpment" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <br />
            <br />
            <br />
            <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClick="btnSubmit_Click" />
            <asp:Label ID="lblError" runat="server" Text="Label"></asp:Label>
        </div>
    </form>
</body>
</html>
