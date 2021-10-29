<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ACR.aspx.cs" Inherits="ACR" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        .Titlebig
        {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: normal;
            text-decoration: none;
        }

        span
        {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }

        input, select
        {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }


        .style3
        {
            width: 11%;
        }


        .style4
        {
            width: 1%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 100%">
                <tr>
                    <td colspan="3">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3" class="Titlebig">Gresham Partners, LLC
                    </td>
                </tr>
                <tr>
                    <td class="Titlebig" colspan="3">Annual Client Review Report</td>
                </tr>
                <tr>
                    <td style="height: 18px" valign="top" colspan="3">
                        <br />
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                        <br />
                    </td>
                </tr>
                <tr>
                    <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblSelect" runat="server"
                            Text="Please select the following fields to generate your report: "></asp:Label>
                        <br />
                    </td>
                </tr>
                <tr>
                    <td class="style3">
                        <asp:Label ID="lblSelectYear" runat="server" Text="Year"></asp:Label></td>
                    <td class="style4" style="color: #FFFFFF">: </td>
                    <td>
                        <asp:DropDownList ID="ddlYear" runat="server" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlYear_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="style3">
                        <asp:Label ID="lblRespoParty" runat="server" Text="Responsible Party"></asp:Label></td>
                    <td class="style4" style="color: #FFFFFF">: </td>
                    <td>
                        <asp:DropDownList ID="ddlRespParty" runat="server" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlRespParty_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="style3">
                        <asp:Label ID="lblHH" runat="server" Text="Household"></asp:Label></td>
                    <td class="style4" style="color: #FFFFFF">: </td>
                    <td>
                        <asp:DropDownList ID="ddlHH" runat="server" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlHH_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="style3">
                        <asp:Label ID="lblTaskType" runat="server" Text="Task Type"></asp:Label></td>
                    <td class="style4" style="color: #FFFFFF">: </td>
                    <td>
                        <asp:DropDownList ID="ddlTaskType" runat="server" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlTaskType_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td valign="top" class="style3">
                        <asp:Label ID="lblStatus" runat="server" Text="Status"></asp:Label></td>
                    <td valign="top" class="style4" style="color: #FFFFFF">: </td>
                    <td valign="top">
                        <asp:ListBox ID="ListBoxStatuses" runat="server" Width="177px"
                            SelectionMode="Multiple" AutoPostBack="True"
                            OnSelectedIndexChanged="ListBoxStatuses_SelectedIndexChanged"></asp:ListBox>
                    </td>
                </tr>
                <tr>
                    <td class="style3">&nbsp;</td>
                    <td class="style4">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="style3">&nbsp;</td>
                    <td class="style4">&nbsp;</td>
                    <td>
                        <asp:Button ID="btnGeneratePDF" runat="server" Text="Generate Report"
                            OnClick="btnGeneratePDF_Click" ToolTip="Click to generate the PDF File"
                            Visible="False" />
                        <asp:Button ID="btnGeneratePDFReport" runat="server"
                            OnClick="btnGeneratePDFReport_Click" Text="Generate PDF Report" />
                        <br />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
