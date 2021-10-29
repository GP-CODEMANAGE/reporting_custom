<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Configuration.aspx.cs" Inherits="Samples_Feature_Configuration_Configuration" %>
<%@ Register TagPrefix="kswc" Namespace="Karamasoft.WebControls.UltimateEditor" Assembly="UltimateEditor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Configuration</title>
	<link href="../../../KaramasoftStyles.css" type="text/css" rel="stylesheet" />
    <link href="KaramasoftStyles.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
		<table class="PageText" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<fieldset><legend>Select Configuration</legend><asp:radiobuttonlist id="rblConfiguration" runat="server" RepeatDirection="Horizontal" CssClass="PageText"
							AutoPostBack="True" RepeatLayout="Flow" OnSelectedIndexChanged="rblConfiguration_SelectedIndexChanged">
							<asp:ListItem Value="Full" Selected="True">Full &nbsp;</asp:ListItem>
							<asp:ListItem Value="Default">Default &nbsp;</asp:ListItem>
							<asp:ListItem Value="Basic">Basic</asp:ListItem>
						</asp:radiobuttonlist></fieldset>
						<br></td>
			</tr>
			<tr>
				<td><kswc:ultimateeditor id="UltimateEditor1" runat="server" EditorSource="~/UltimateEditorInclude/UltimateEditorFull.xml"
						DisplayCharCount="True" MaxCharCount="50000" DisplayWordCount="True" MaxWordCount="10000" Resizable="True">
					</kswc:ultimateeditor></td>
			</tr>
			<tr>
				<td>
					<br>
					<fieldset><legend>Description</legend><div>
						You can easily configure your editor by setting the <b>EditorSource</b> property 
						to a toolbar XML file.<br>
						<br>
						UltimateEditor provides three built-in toolbar XML files:<br>
						<ul>
							<li>
								<b>UltimateEditor.xml</b>
							(default)
							<li>
								<b>UltimateEditorBasic.xml</b>
							<li>
								<b>UltimateEditorFull.xml</b></li></ul>
					</div></fieldset>
				</td>
			</tr>
		</table>
    </form>
</body>
</html>
