<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmTNRAgendaRpt.aspx.cs" Inherits="frmTNRAgendaRpt" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TNR Agenda Report</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>
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

        .auto-style1
        {
            width: 182px;
        }
    </style>

    <script>
        function ValidateDate(source, args) {

            var txtPriorDate = document.getElementById("txtStartDate").value;
            var txtEndDate = document.getElementById("txtendDate").value;

            if (txtPriorDate != "") {
                var txtDate = new Date(txtPriorDate);
                var txtToDate = new Date(txtEndDate);

                if (txtDate <= txtToDate)
                    args.IsValid = true;
                else {
                    args.IsValid = false;
                    alert("Please select greater date than Start Date")
                }
            }
        }

        function showexcrpt() {

           // document.getElementById("trexception").style.display = "";
            var labelObj = document.getElementById("<%= lblMessage.ClientID %>");
            labelObj.value = "";
            // alert('g');
        }



        function ShowAlert() {

        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table width="100%">
                <tr>
                    <td colspan="2">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="Titlebig">Gresham Partners, LLC
                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <td class="Titlebig" colspan="2">TNR Agenda Report</td>
                </tr>
                <tr>
                    <td style="height: 18px" valign="top" colspan="2">
                        <br />
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                        <br />
                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <td class="auto-style1">
                        <asp:Label ID="Label2" runat="server" Text="Start Date"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="txtStartDate" runat="server" Width="119px"></asp:TextBox>
                        <a onclick="showCalendarControl(txtStartDate)">
                            <img id="img2" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>
                        <asp:RegularExpressionValidator
                            ID="RegularExpressionValidator2" runat="server" ErrorMessage="Invalid Start Date"
                            ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
                            ControlToValidate="txtStartDate" Display="Dynamic">Please enter correct date</asp:RegularExpressionValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" ControlToValidate="txtStartDate" runat="server" Display="Dynamic" ErrorMessage="Please Select Start Date"></asp:RequiredFieldValidator>
                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <td class="auto-style1">
                        <asp:Label ID="Label3" runat="server" Text="End Date"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="txtendDate" runat="server" Width="119px"></asp:TextBox>
                        <a onclick="showCalendarControl(txtendDate)">
                            <img id="img3" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="Invalid End Date"
                            ControlToValidate="txtendDate" ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
                            Display="Dynamic">*</asp:RegularExpressionValidator>
                        <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="Please select greater date than Start Date"
                            ClientValidationFunction="ValidateDate" Display="Dynamic">Please enter correct date</asp:CustomValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" ControlToValidate="txtendDate" runat="server" Display="Dynamic" ErrorMessage="Please Select End Date"></asp:RequiredFieldValidator>
                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblFund" runat="server" Text="Fund Group"></asp:Label></td>
                    <td>
                        <%--<asp:DropDownList ID="ddlFund" runat="server"></asp:DropDownList>--%>
                        <asp:ListBox ID="lstFund" runat="server" onchange="ClearLabel();" SelectionMode="Multiple"
                                    Rows="5"></asp:ListBox>
                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <td class="auto-style1">Report Type</td>
                    <td>
                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" OnSelectedIndexChanged="RadioButtonList1_SelectedIndexChanged" AutoPostBack="True">
                            <asp:ListItem Text="PDF" Enabled="true" Selected="True" Value="1"></asp:ListItem>
                            <asp:ListItem Text="Excel" Value="2"></asp:ListItem>
                        </asp:RadioButtonList>

                    </td>
                    <td class="style2"></td>
                </tr>
                <tr>
                    <%-- <td colspan="3">
                        <asp:LinkButton ID="lnkPCAReport" runat="server" OnClick="lnkPCAReport_Click" Visible="false">PCA Report</asp:LinkButton><br />
                        <asp:LinkButton ID="lnkPCAReport2" runat="server" OnClick="lnkPCAReport2_Click" Visible="false">PCA Report 2</asp:LinkButton><br />
                    </td>--%>
                </tr>
                <tr id="trexception" runat="server">
                    <td>&nbsp;</td>
                    <td colspan="2">
                        <asp:LinkButton ID="lnkException" runat="server" OnClick="lnkException_Click">Exception Report</asp:LinkButton>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                </tr>
                <tr>

                    <td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnGeneratePDF" runat="server" Text="Generate Report"
                            OnClick="btnGeneratePDF_Click" ToolTip="Click to generate the PDF File" OnClientClick="javascript:showexcrpt();" />
                        <%--    <asp:Button ID="btnHidden" runat="server" Text="Button" />--%>
                        <%--OnClientClick="javascript:showexcrpt();"--%>
                    </td>
                    <td class="style2"></td>
                </tr>
            </table>
        </div>

    </form>
</body>
</html>
