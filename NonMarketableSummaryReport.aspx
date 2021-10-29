<%@ Page Language="C#" AutoEventWireup="true" CodeFile="NonMarketableSummaryReport.aspx.cs"
    Inherits="NonMarketableSummaryReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Non-Marketable Summary Report</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <style type="text/css">
    .CellTopBorder
    {
	    border-top-color:Gray; border-top:solid; border-top-width:thick;
    }
    
      
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table>
                <tr>
                    <td align="left" style="width: 204px; height: 29px">
                        <img src="images/Gresham_Logo__.jpg" />&nbsp;
                    </td>
                    <td style="height: 29px">
                    </td>
                    <td style="height: 29px">
                    </td>
                </tr>
                <tr>
                    <td align="left" style="width: 204px; height: 29px">
                        Gresham Partners, LLC
                    </td>
                    <td style="height: 29px">
                    </td>
                    <td style="height: 29px">
                    </td>
                </tr>
                <tr>
                    <td align="left" style="width: 204px; height: 29px">
                        <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Medium"
                            Text="Non-Marketable Summary Report" Width="313px"></asp:Label></td>
                    <td style="height: 29px">
                    </td>
                    <td style="height: 29px">
                    </td>
                </tr>
                <tr>
                    <td style="width: 204px;">
                        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <table runat="server" id="tblFilter" style="width: 496px">
                            <tr>
                                <td style="width: 182px; height: 88px">
                                    <strong>Type</strong></td>
                                <td style="height: 88px">
                                    <asp:ListBox ID="lstType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="lstType_SelectedIndexChanged"
                                        SelectionMode="Multiple"></asp:ListBox></td>
                                <td style="width: 3px; height: 88px">
                                </td>
                                <td style="width: 4px; height: 88px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 182px; height: 88px">
                                    <strong>Partnership</strong></td>
                                <td style="height: 88px">
                                    <asp:ListBox ID="lstbxPartnership" runat="server" Height="143px" Rows="12"
                                        SelectionMode="Multiple"></asp:ListBox></td>
                                <td style="width: 3px; height: 88px">
                                </td>
                                <td style="width: 4px; height: 88px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 182px">
                                    <strong>As Of Date</strong></td>
                                <td>
                                    <asp:TextBox ID="txtAsOfDate" runat="server"></asp:TextBox><a onclick="showCalendarControl(txtAsOfDate)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" /></a></td>
                                <td style="width: 3px">
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 182px">
                                </td>
                                <td>
                                </td>
                                <td style="width: 3px">
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 182px">
                                </td>
                                <td align="left">
                                    <asp:Button ID="btnExportToExcel1" runat="server" OnClick="btnExportToExcel_Click"
                                        Text="Export To Excel" />
                                </td>
                                <td style="width: 3px">
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
