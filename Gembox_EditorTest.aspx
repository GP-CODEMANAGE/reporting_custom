<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Gembox_EditorTest.aspx.cs" Inherits="Gembox_EditorTest" %>
<%@ Register TagPrefix="cc" Namespace="Winthusiasm.HtmlEditor" Assembly="Winthusiasm.HtmlEditor" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <FCKeditorV2:FCKeditor ID="txtFundDesc1" runat="server" BasePath="~/FCKeditor/" Height="250px" SkinPath="skins/silver/" ToolbarSet="Minimal" Width="60%">
            </FCKeditorV2:FCKeditor>
            <br />
            <br />
            <asp:Label ID="Label1" runat="server" Text="Label" Visible="False"></asp:Label>
        </div>
        <asp:Button ID="Inser_CRM" runat="server" OnClick="Inser_CRM_Click" Text="Insert into CRM" />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Create PDF" />
    </form>
</body>
</html>
