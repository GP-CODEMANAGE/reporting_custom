<%@ Page Language="C#" AutoEventWireup="true" CodeFile="TestPage.aspx.cs" Inherits="TestPage" %>
<%@ Register TagPrefix="cc" Namespace="Winthusiasm.HtmlEditor" Assembly="Winthusiasm.HtmlEditor" %>

<%@ Register TagPrefix="kswc" Namespace="Karamasoft.WebControls.UltimateEditor" Assembly="UltimateEditor" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link id="style1" href="common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="common/Calendar.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <FCKeditorV2:FCKeditor ID="txtFundDesc1" runat="server" BasePath="~/FCKeditor/" Height="250px" SkinPath="skins/silver/" ToolbarSet="Minimal" Width="60%">
            </FCKeditorV2:FCKeditor>
            <br />
            <br />
           <kswc:ultimateeditor id="Ultimateeditor1" runat="server" EditorSource="~/UltimateEditorInclude/UltimateEditorFull.xml"
						DisplayCharCount="True" MaxCharCount="50000" DisplayWordCount="True" MaxWordCount="10000" Resizable="True">
					</kswc:ultimateeditor>
            <br />
            <br />
            <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Button" />
        </div>
    </form>
</body>
</html>
