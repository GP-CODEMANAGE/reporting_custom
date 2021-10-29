<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BatchReportTEST.aspx.cs" Inherits="BatchReportTEST" Culture="en-US" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Advent Batch Report (NEW)</title>
    <link href="../common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="../common/Calendar.js" type="text/javascript"></script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 839px">
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:Label ID="lblError" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 18px">
                        <table>
                              <tr>
                                <td>Household Type:
                                </td>
                                <td style="width: 220px">&nbsp;<asp:DropDownList ID="ddlHouseHoldType" runat="server" AutoPostBack="true" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" >
                                </asp:DropDownList>
                                </td>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td>Household:
                                </td>
                                <td style="width: 220px">&nbsp;<asp:DropDownList ID="ddlHouseHold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHouseHold_SelectedIndexChanged">
                                </asp:DropDownList>
                                </td>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td>Batch Type:</td>
                                <td style="width: 220px">&nbsp;<asp:DropDownList ID="ddlBatchType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlBatchType_SelectedIndexChanged">
                                    <asp:ListItem Selected="True" Value="0">Please Select</asp:ListItem>
                                    <asp:ListItem Value="1">MTGBK</asp:ListItem>
                                    <asp:ListItem Value="2">Q</asp:ListItem>
                                    <asp:ListItem Value="3">M</asp:ListItem>
                                </asp:DropDownList>
                                </td>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td>Prior Date:</td>
                                <td style="width: 220px">&nbsp;<asp:TextBox ID="txtPriorDate" runat="server" ValidationGroup="myGroup"></asp:TextBox>&nbsp;
                                    <a onclick="showCalendarControl(txtPriorDate)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" runat="server" />&nbsp;
                                    </a>
                                </td>
                                <td>&nbsp;<asp:CheckBox ID="chkNoComparison" runat="server" Text="no comparison line"
                                    AutoPostBack="True" OnCheckedChanged="chkNoComparison_CheckedChanged" Font-Bold="True" /></td>
                                <td>&nbsp;<asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server"
                                    ErrorMessage="Invalid Prior Date" ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
                                    ControlToValidate="txtPriorDate" Display="Static" ValidationGroup="myGroup">*</asp:RegularExpressionValidator></td>
                            </tr>
                            <tr>
                                <td style="height: 26px">As of &nbsp;Date:</td>
                                <td style="height: 26px; width: 220px;">&nbsp;<asp:TextBox ID="txtEndDate" runat="server"></asp:TextBox>&nbsp;<a onclick="showCalendarControl(txtEndDate)">
                                    <img id="Img2" alt="" border="0" src="images/calander.png" /></a>
                                </td>
                                <td style="height: 26px">
                                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="Invalid End Date"
                                        ControlToValidate="txtEndDate" ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$">*</asp:RegularExpressionValidator></td>
                                <td style="height: 26px"></td>
                            </tr>
                            <tr>
                                <td style="height: 26px">Suppress Manager Detail:&nbsp;
                                </td>
                                <td style="width: 220px; height: 26px">
                                    <asp:CheckBox ID="chkSuppressManagerDetail" runat="server" /></td>
                                <td style="height: 26px"></td>
                                <td style="height: 26px"></td>
                            </tr>
                            <tr>
                                <td style="height: 26px"></td>
                                <td style="width: 220px; height: 26px"></td>
                                <td style="height: 26px"></td>
                                <td style="height: 26px"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="False" TabIndex="1"
                            ToolTip="Batch List" Width="100%" OnRowDataBound="gvList_RowDataBound">
                            <Columns>
                                <asp:BoundField DataField="Ssi_batchId" HeaderText="Ssi_batchId" Visible="False" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" ID="chkbSelectBatch" Checked="false" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField HeaderText="Batch Name" DataField="BatchName" SortExpression="BatchName" />
                                <asp:BoundField HeaderText="Contact" DataField="Ssi_ContactIdName" SortExpression="Ssi_ContactIdName" />
                                <asp:BoundField HeaderText="Created By" DataField="CreatedByName" SortExpression="CreatedByName" />
                                <asp:BoundField DataField="FolderNameTxt" HeaderText="FolderName" Visible="False" />
                                <asp:BoundField DataField="HouseholdNameTxt" HeaderText="HouseholdNameTxt" Visible="False" />
                                <asp:BoundField DataField="PdfFileName" HeaderText="PdfFileName" Visible="False" />
                                <%-- <asp:BoundField HeaderText="" DataField="" SortExpression="" />
                                <asp:BoundField DataField="" DataFormatString="{0:f3}" HeaderText="" HtmlEncode="False" />
                                <asp:BoundField DataField="" HeaderText="" DataFormatString="{0:f3}" HtmlEncode="False" />--%>
                            </Columns>
                            <HeaderStyle Height="10px" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td align="right">
                        <asp:Button ID="btnGenerateReport" runat="server" OnClientClick="return hidenme(this);"
                            OnClick="btnGenerateReport_Click" Text="Generate Report" />
                        <div id="divdot" style="display: none;">
                            ....
                        </div>

                        <script type="text/javascript">

                            function hidenme(obj) {
                                var isValid = Page_ClientValidate('');
                                if (isValid) {
                                    obj.style.display = "none";
                                    document.getElementById("divdot").style.display = "";
                                    return true;
                                }
                                else {
                                    return false;
                                }
                            }



                        </script>

                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List"
                            ShowMessageBox="True" ShowSummary="False" />
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
            </table>
        </div>
    </form>

    <script language="javascript" type="text/javascript">
        function ClearLabel() {
            document.getElementById('<%= lblError.ClientID%>').innerHTML = "";
    }

    </script>

</body>
</html>
