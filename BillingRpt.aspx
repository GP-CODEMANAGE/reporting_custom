<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BillingRpt.aspx.cs" Inherits="BillingRpt" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Billing Fee Calc and Worksheet Report</title>
    <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table class="auto-style34">

                <tr>
                    <td colspan="3" class="auto-style35">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                    <td class="auto-style63">&nbsp;</td>
                    <td class="auto-style36">&nbsp;</td>
                    <td class="auto-style37">&nbsp;</td>
                    <td class="auto-style38">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3" class="auto-style39">Gresham Partners, LLC
                    </td>
                    <td class="auto-style64">&nbsp;</td>
                    <td class="auto-style40">&nbsp;</td>
                    <td class="auto-style41">&nbsp;</td>
                    <td class="auto-style42">&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style39" colspan="3">Billing Fee Calc and Worksheet Report
                    </td>
                    <td class="auto-style64">&nbsp;</td>
                    <td class="auto-style40">&nbsp;</td>
                    <td class="auto-style41">&nbsp;</td>
                    <td class="auto-style42">&nbsp;</td>
                </tr>
                <tr>
                    <td valign="top" colspan="7" class="auto-style43">
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ValidationGroup="vgBIE" ShowMessageBox="true" DisplayMode="List" ShowSummary="false" />
                    </td>
                </tr>
                <tr>
                    <td valign="top" colspan="6" class="auto-style45">
                        <asp:Label ID="lblSelect" runat="server"
                            Text="Please select the following fields to generate your invoice: "></asp:Label>
                        <br />
                    </td>
                    <td valign="top" class="auto-style30">&nbsp;</td>
                </tr>

                <tr>

                    <td class="auto-style46">
                        <asp:Label ID="lblSelectAdv" runat="server" Text="Advisor"></asp:Label></td>
                    <td class="auto-style47" style="color: #FFFFFF">: </td>
                    <td class="auto-style48">
                        <asp:DropDownList ID="ddlAdvisor" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAdvisor_SelectedIndexChanged" TabIndex="1">
                        </asp:DropDownList>
                        <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlAdvisor" Display="Dynamic" ErrorMessage="Select Advisor" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%>
                    </td>
                    <td class="auto-style47" style="width: 30px">&nbsp;</td>
                    <td class="auto-style49">&nbsp;</td>
                    <td class="auto-style50">&nbsp;</td>
                    <td class="auto-style51">&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style46">
                        <asp:Label ID="lblHH" runat="server" Text="Household"></asp:Label></td>
                    <td class="auto-style47" style="color: #FFFFFF">: </td>
                    <td class="auto-style48">
                        <asp:DropDownList ID="ddlHH" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHH_SelectedIndexChanged" TabIndex="2">
                        </asp:DropDownList>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="ddlHH" Display="Dynamic" ErrorMessage="Select Household" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>
                    <td valign="top" class="auto-style47">&nbsp;</td>
                    <td valign="top" class="auto-style49">&nbsp;</td>

                    <td class="auto-style51">&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style46">
                        <asp:Label ID="lblBilFor" runat="server" Text="Billing For"></asp:Label></td>
                    <td class="auto-style47" style="color: #FFFFFF">: </td>
                    <td class="auto-style48">
                        <asp:DropDownList ID="ddlBillFor" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlBillFor_SelectedIndexChanged1" TabIndex="3">
                        </asp:DropDownList>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlBillFor" Display="Dynamic" ErrorMessage="Select Billing For" InitialValue="ALL" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>
                    <td valign="top" class="auto-style47">&nbsp;</td>
                    <td valign="top" class="auto-style49">&nbsp;</td>
                    <td valign="top" class="auto-style50">&nbsp;</td>
                    <%-- <td class="auto-style51">
                        <a onclick="showCalendarControl(txtBillingPeriod)">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="txtBillingPeriod" Display="Dynamic" ErrorMessage="Select Billing period" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>--%>
                </tr>
                <tr>
                    <td class="auto-style46">
                        <asp:Label ID="lblAumAsofDate" runat="server" Text="AUM as of Date"></asp:Label></td>
                    <td class="auto-style47" style="color: #FFFFFF"></td>
                    <td class="auto-style48">
                        <asp:TextBox ID="txtAUMDate" runat="server" Width="119px" onChange="selectMonths(this.value);" AutoPostBack="True" OnTextChanged="txtAUMDate_TextChanged" TabIndex="4"></asp:TextBox>
                        <a onclick="showCalendarControl(txtAUMDate)">
                            <img id="img2" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="txtAUMDate" Display="Dynamic" ErrorMessage="Select AUM as of Date" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>

                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td> </td>
                     <td>  <asp:ListBox ID="lbGroup" runat="server" Width="391px"  Visible="false"></asp:ListBox></td>
                    </tr>
                <tr>
                    <td class="auto-style46">&nbsp;</td>
                    <td class="auto-style47">&nbsp;</td>
                    <td class="auto-style48">
                        <asp:Button ID="btnSubmit" runat="server" Text="Generate Report" ToolTip="Click to generate the report" ValidationGroup="vgBIE" OnClick="btnSubmit_Click" TabIndex="8" />
                        <br />
                    </td>
                    <td class="auto-style47">&nbsp;</td>
                    <td class="auto-style49">&nbsp;</td>
                    <td class="auto-style50">&nbsp;</td>
                    <td class="auto-style51">&nbsp;</td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
