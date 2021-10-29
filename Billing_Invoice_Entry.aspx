<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Billing_Invoice_Entry.aspx.cs" Inherits="Billing_Invoice_Entry" Culture="en-US" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Billing Invoice Entry</title>
    <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>


    <style type="text/css">
        .TextAlgRgh {
            text-align: right;
        }

        .auto-style1 {
            height: 32px;
        }
    </style>
    <%--<style type="text/css">
        .auto-style30
        {
            height: 19px;
            width: 3%;
        }

        .auto-style34
        {
            width: 75%;
            height: 494px;
        }

        .auto-style35
        {
            height: 25px;
        }

        .auto-style36
        {
            width: 19%;
            height: 25px;
        }

        .auto-style37
        {
            width: 6%;
            height: 25px;
        }

        .auto-style38
        {
            width: 3%;
            height: 25px;
        }

        .auto-style39
        {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: normal;
            text-decoration: none;
            height: 21px;
        }

        .auto-style40
        {
            width: 19%;
            height: 21px;
        }

        .auto-style41
        {
            width: 6%;
            height: 21px;
        }

        .auto-style42
        {
            width: 3%;
            height: 21px;
        }

        .auto-style43
        {
            height: auto;
        }

        .auto-style45
        {
            height: 19px;
        }

        .auto-style46
        {
            width: 25%;
            height: 28px;
        }

        .auto-style47
        {
            width: 1%;
            height: 28px;
        }

        .auto-style48
        {
            width: 164px;
            height: 28px;
        }

        .auto-style49
        {
            width: 19%;
            height: 28px;
        }

        .auto-style50
        {
            width: 6%;
            height: 28px;
        }

        .auto-style51
        {
            width: 3%;
            height: 28px;
        }

        .auto-style52
        {
            width: 25%;
            height: 19px;
        }

        .auto-style53
        {
            width: 1%;
            height: 19px;
        }

        .auto-style54
        {
            width: 164px;
            height: 19px;
        }

        .auto-style55
        {
            width: 19%;
            height: 19px;
        }

        .auto-style56
        {
            width: 6%;
            height: 19px;
        }

        .auto-style57
        {
            height: 29px;
            width: 25%;
        }

        .auto-style58
        {
            height: 29px;
            width: 1%;
        }

        .auto-style59
        {
            height: 29px;
            width: 164px;
        }

        .auto-style60
        {
            height: 29px;
            width: 19%;
        }

        .auto-style61
        {
            height: 29px;
            width: 6%;
        }

        .auto-style62
        {
            height: 29px;
            width: 3%;
        }

        .auto-style63
        {
            width: 1%;
            height: 25px;
        }

        .auto-style64
        {
            width: 1%;
            height: 21px;
        }
    </style>--%>

    <script language="Javascript">
        /**
         * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
         */
        // Declaring valid date character, minimum year and maximum year
        var dtCh = "/";
        var minYear = 1900;
        var maxYear = 2100;

        function isInteger(s) {
            var i;
            for (i = 0; i < s.length; i++) {
                // Check that current character is number.
                var c = s.charAt(i);
                if (((c < "0") || (c > "9"))) return false;
            }
            // All characters are numbers.
            return true;
        }

        function stripCharsInBag(s, bag) {
            var i;
            var returnString = "";
            // Search through string's characters one by one.
            // If character is not in bag, append to returnString.
            for (i = 0; i < s.length; i++) {
                var c = s.charAt(i);
                if (bag.indexOf(c) == -1) returnString += c;
            }
            return returnString;
        }

        function daysInFebruary(year) {
            // February has 29 days in any year evenly divisible by four,
            // EXCEPT for centurial years which are not also divisible by 400.
            return (((year % 4 == 0) && ((!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28);
        }
        function DaysArray(n) {
            for (var i = 1; i <= n; i++) {
                this[i] = 31
                if (i == 4 || i == 6 || i == 9 || i == 11) { this[i] = 30 }
                if (i == 2) { this[i] = 29 }
            }
            return this
        }

        function isDate(dtStr) {
            var daysInMonth = DaysArray(12)
            var pos1 = dtStr.indexOf(dtCh)
            var pos2 = dtStr.indexOf(dtCh, pos1 + 1)
            var strMonth = dtStr.substring(0, pos1)
            var strDay = dtStr.substring(pos1 + 1, pos2)
            var strYear = dtStr.substring(pos2 + 1)
            strYr = strYear
            if (strDay.charAt(0) == "0" && strDay.length > 1) strDay = strDay.substring(1)
            if (strMonth.charAt(0) == "0" && strMonth.length > 1) strMonth = strMonth.substring(1)
            for (var i = 1; i <= 3; i++) {
                if (strYr.charAt(0) == "0" && strYr.length > 1) strYr = strYr.substring(1)
            }
            month = parseInt(strMonth)
            day = parseInt(strDay)
            year = parseInt(strYr)
            if (pos1 == -1 || pos2 == -1) {

                return false
            }
            if (strMonth.length < 1 || month < 1 || month > 12) {

                return false
            }
            if (strDay.length < 1 || day < 1 || day > 31 || (month == 2 && day > daysInFebruary(year)) || day > daysInMonth[month]) {

                return false
            }
            if (strYear.length != 4 || year == 0 || year < minYear || year > maxYear) {

                return false
            }
            if (dtStr.indexOf(dtCh, pos2 + 1) != -1 || isInteger(stripCharsInBag(dtStr, dtCh)) == false) {

                return false
            }
            return true
        }

        function ValidateForm(oSrc, args) {

            args.IsValid = isDate(args.Value);

        }

        function selectMonths(posDate) {
            // alert(posDate);
            if (posDate != "") {
                var dt = new Date(posDate);
                var month = dt.getMonth() + 1;
                if (month == 1 || month == 2 || month == 3) {
                    document.getElementById('ddlMonths').value = 2;
                }
                else if (month == 4 || month == 5 || month == 6) {
                    document.getElementById('ddlMonths').value = 3;
                }
                else if (month == 7 || month == 8 || month == 9) {
                    document.getElementById('ddlMonths').value = 4;
                }
                else if (month == 10 || month == 11 || month == 12) {
                    document.getElementById('ddlMonths').value = 1;
                }
            }

        }


        function Confirm() {
            var confirm_value = document.createElement("INPUT");
            confirm_value.type = "hidden";
            confirm_value.name = "confirm_value";
            if (confirm("An Invoice already exists for the selected household billing entity, press ok to overwrite the invoice or cancel keep the original invoice")) {
                confirm_value.value = "Yes";
            } else {
                confirm_value.value = "No";
            }
            document.forms[0].appendChild(confirm_value);
        }


        function isNumberKey(sender, evt) {
            var txt = sender.value;
            var dotcontainer = txt.split('.');
            var charCode = (evt.which) ? evt.which : event.keyCode;

            if (!(dotcontainer.length == 1 && charCode == 46) && charCode > 31 && (charCode < 48 || charCode > 57) && charCode != 45)
                return false;

            return true;
        }

        function showAdjustmentSection() {
            var trAdjustmentAmt = document.getElementById("trAdjustmentAmt");
            var trAdjustmentReason = document.getElementById("trAdjustmentReason");
            var trAdjQtrFee = document.getElementById("trAdjQtrFee");
            var hidden = document.getElementById('<%= Hcheckadjvalue.ClientID %>');

            if (hidden != null) {
                hidden.value = "1";

            }

            if (hidden.value == "1") {
                trAdjustmentAmt.style.display = '';
                trAdjustmentReason.style.display = '';
                trAdjQtrFee.style.display = '';
            }
        }

        function showAdjustmentSectionOnLoad() {
            var trAdjustmentAmt = document.getElementById("trAdjustmentAmt");
            var trAdjustmentReason = document.getElementById("trAdjustmentReason");
            var trAdjQtrFee = document.getElementById("trAdjQtrFee");
            var hidden = document.getElementById('<%= Hcheckadjvalue.ClientID %>');
            var txtadjamt = document.getElementById('<%= txtAdjAmt.ClientID %>');
            var txtAdjReason = document.getElementById('<%= txtAdjReason.ClientID %>');
            var txtAdjQtrFee = document.getElementById('<%= txtAdjQtrFee.ClientID %>');

            /* if Adjustment amount is there show 3 adjustment controls section */
            if (txtadjamt.value != null && txtadjamt.value != '') {
                hidden.value = '1';
            }
            else {
                hidden.value = '';
            }

            // alert(hidden.value);
            if (hidden.value == "1") {
                trAdjustmentAmt.style.display = '';
                trAdjustmentReason.style.display = '';
                trAdjQtrFee.style.display = '';
            }
            else {
                trAdjustmentAmt.style.display = 'none';
                trAdjustmentReason.style.display = 'none';
                trAdjQtrFee.style.display = 'none';
                txtAdjReason.value = '';
                txtadjamt.value = '';
                txtAdjQtrFee.value = '';

            }
        }

        // function hideAdjustmentSection() {



    </script>
</head>
<body onload="showAdjustmentSectionOnLoad();return false;">
    <form id="form1" runat="server">
        <div>
            <table class="auto-style34">

                <tr>
                    <td colspan="3" class="auto-style35">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>

                </tr>
                <tr>
                    <td colspan="3" class="auto-style39">Gresham Partners, LLC
                    </td>

                </tr>
                <tr>
                    <td class="auto-style39" colspan="3">Billing Invoice Entry
                    </td>

                </tr>
                <tr>
                    <td valign="top" colspan="7" class="auto-style43">
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                        <br />

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

                    <td rowspan="2">
                        <asp:Label ID="Label3" runat="server" Text="Notes" Font-Bold="true"></asp:Label>
                    </td>
                    <td rowspan="2" colspan="3">
                        <asp:TextBox ID="txtNotes" runat="server" Height="52px" TextMode="MultiLine"
                            Width="374px"></asp:TextBox>
                        &nbsp;</td>


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

                    <td></td>
                    <td>
                        <asp:Button ID="btnDeleteAndCal" runat="server"
                            Text="Delete Invoice & Re-Calculate" Visible="false"
                            Width="227px" OnClick="btnDeleteAndCal_Click" />
                    </td>



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

                        &nbsp;&nbsp; &nbsp;
                    </td>

                    <td>
                        <asp:Label ID="lblBillingPeriod" runat="server" Text="Billing Period" Font-Bold="true"></asp:Label></td>

                    <td>&nbsp;<asp:DropDownList ID="ddlMonths" runat="server" BackColor="#C7EDF7">
                        <asp:ListItem Value="1">February – April</asp:ListItem>
                        <asp:ListItem Value="2">May – July</asp:ListItem>
                        <asp:ListItem Value="3">August – October</asp:ListItem>
                        <asp:ListItem Value="4">November – January </asp:ListItem>
                    </asp:DropDownList>
                    </td>

                </tr>
                <tr>
                    <td valign="top" class="auto-style47">&nbsp;</td>
                    <td valign="top" class="auto-style50">&nbsp;</td>
                    <td valign="top" class="auto-style47" colspan="3">

                        <asp:ListBox ID="lbGroup" runat="server" Width="391px" Visible="false"></asp:ListBox>

                    </td>

                </tr>

                <tr>
                    <td class="auto-style46">Fee Schedule
                    </td>
                    <td class="auto-style47" style="color: #FFFFFF">: </td>
                    <td>

                        <%--     <asp:RadioButtonList ID="rdolstClientType" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" OnSelectedIndexChanged="rdolstClientType_SelectedIndexChanged">
                            <asp:ListItem Text="New" Value="New" Selected></asp:ListItem>
                            <asp:ListItem Text="Existing" Value="Existing"></asp:ListItem>                            
                        </asp:RadioButtonList>--%>
                        <asp:DropDownList ID="ddlClientType" runat="server" AutoPostBack="true" Width="385px" BackColor="#C7EDF7" OnSelectedIndexChanged="ddlClientType_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" ControlToValidate="ddlClientType" Display="Dynamic" ErrorMessage="Select Fee Schedule" InitialValue="0" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>
                    <td class="auto-style46" colspan="2">
                        <asp:Label ID="lblbpsfees" runat="server" Text="Variable Discount on first 25 mil (bps)" Visible="False"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="txtbpsfee" runat="server" OnKeyPress="return isNumberKey(this, event);" CssClass="TextAlgRgh" Visible="False"></asp:TextBox>
                    </td>
                </tr>

                <tr>
                    <td valign="top" class="auto-style46">
                        <br />
                        <asp:Label ID="lblTotAUM" runat="server" Text="Total AUM"></asp:Label></td>
                    <td valign="top" class="auto-style47" style="color: #FFFFFF">: </td>
                    <td valign="top" class="auto-style48" c>

                        <asp:TextBox ID="txtTotalAUM" runat="server" AutoPostBack="True" OnKeyPress="return isNumberKey(this, event);" OnTextChanged="txtTotalAUM_TextChanged" TabIndex="5" CssClass="TextAlgRgh"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="txtTotalAUM" Display="Dynamic" ErrorMessage="Enter Total AUM" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                        &nbsp;
                        <asp:LinkButton ID="LinkButton1" runat="server" OnClick="LinkButton1_Click" Visible="false">Worksheet PDF</asp:LinkButton>

                    </td>
                    <td class="auto-style46" colspan="2">
                        <asp:Label ID="lblMinValu" runat="server" Text="Minimum fee as $: <br /> (If blank, then comparison will be 180K)" Visible="False"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="txtMinVal" runat="server" AutoPostBack="False" OnKeyPress="return isNumberKey(this, event);" CssClass="TextAlgRgh" Visible="False" OnTextChanged="txtMinVal_TextChanged"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style57">
                        <br />
                        <asp:Label ID="lblBillAUM" runat="server" Text="Total Billable Assets"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">


                        <asp:TextBox ID="txtTotalBillingAUM" runat="server" AutoPostBack="True" CssClass="TextAlgRgh" OnKeyPress="return isNumberKey(this, event);" TabIndex="6" OnTextChanged="txtTotalBillingAUM_TextChanged"></asp:TextBox>


                    </td>




                    <td class="auto-style46" colspan="2">
                        <asp:Label ID="lblMaxVal" runat="server" Text="Maximum fee as % of AUM: <br /> (If blank, no maximum comparison will be made) " Visible="False"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="txtMaxVal" runat="server" AutoPostBack="False"
                            OnKeyPress="return isNumberKey(this, event);"
                            CssClass="TextAlgRgh" Visible="False" OnTextChanged="txtMaxValue_TextChanged"></asp:TextBox>
                    </td>

                </tr>


                  <tr>
                    <td class="auto-style57">
                        <br />
                         <asp:Label ID="lblStandardFeeAssets" runat="server" Text="Billing AUM" Visible="False" ></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">
                        <asp:TextBox ID="txtBillAUM" runat="server" OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7" CssClass="TextAlgRgh"></asp:TextBox>

                    </td>

                </tr>
                

                <tr>
                    <%--id="trStdAnnualFeeCalc" runat="server"--%>
                    <td class="auto-style57">
                        <asp:Label ID="lblStdAnnFeeCalc" runat="server" Text="Annual Fee Calc"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style57">
                        <table>
                            <tr>
                                <td>
                                    <asp:TextBox ID="txtStdAnnualFeeCalc" runat="server" OnKeyPress="return isNumberKey(this, event);" OnTextChanged="txtStdAnnualFeeCalc_TextChanged" BackColor="#C7EDF7" CssClass="TextAlgRgh"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="txtStdAnnualFeeCalc" Display="Dynamic" ErrorMessage="Enter annual fee calc" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                                </td>

                               

                            </tr>
                           
                        </table>
                    </td>
                    <td>
                        <asp:Button ID="btnStandardCalculate" runat="server" Text="Calculate/Clear" Visible="False" OnClick="btnStandardCalculate_Click" />
                    </td>


                </tr>
                <%--  <tr>
                    <td></td>
                    <td></td>
                    <td class="auto-style61"></td>

                  
                    <td></td>

                    <td class="auto-style57" colspan="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>

                </tr>--%>
                <tr>
                    <td class="auto-style57" style="width:150px;">
                        <asp:Label  ID="Label5" runat="server" Text="Fee for Administrative Assets" Visible="false"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>

                    <td class="auto-style58">
                        <table>
                            <tr>
                                <td>
                                    <asp:TextBox ID="txtSecurityFee" runat="server" AutoPostBack="True" OnKeyPress="return isNumberKey(this, event);" Enabled="false" Visible="false" CssClass="TextAlgRgh"></asp:TextBox>
                                </td>

                                <td>
                                    &nbsp;&nbsp;    <asp:TextBox ID="txtFeeAUM" runat="server" AutoPostBack="True" OnKeyPress="return isNumberKey(this, event);" Enabled="false" Visible="false" CssClass="TextAlgRgh"></asp:TextBox>
                                </td>

                            </tr>
                            <tr>
                                <td></td>
                                <td>
                                    &nbsp;&nbsp;   <asp:Label ID="lblAssetsUnderAdministration" runat="server" Text="Administrative Assets" Visible="False" ></asp:Label>

                                </td>
                            </tr>
                        </table>
                    </td>

                </tr>
                <tr>
                    <td colspan="3">
                        <asp:GridView ID="gvFlatFee" runat="server" AutoGenerateColumns="False" ShowHeader="False" GridLines="None" Width="450px">
                            <Columns>
                                <%--   <asp:BoundField DataField="Colname" HeaderText=""  >  <HeaderStyle Width="200px" />  <%-- <ItemStyle Width="500px"></ItemStyle>    </asp:BoundField>  --%>

                                <asp:TemplateField ControlStyle-Width="102px">
                                    <ItemTemplate>
                                        <asp:Label ID="lblFlatFee" Width="90px" runat="server" AutoPostBack="true" Text='<%# Bind("ssi_Comments")%>'></asp:Label>

                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField ControlStyle-Width="180px">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtFlatFee" runat="server" Width="180px" AutoPostBack="true" Text='<%# Bind("Fees") %>' DataFormatString="{0:C2}" BackColor="#C7EDF7" CssClass="TextAlgRgh" Enabled="false"></asp:TextBox>

                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </td>
                </tr>


                <%--   <asp:BoundField DataField="Colname" HeaderText=""  >  <HeaderStyle Width="200px" />  <%-- <ItemStyle Width="500px"></ItemStyle>    </asp:BoundField>  --%>



                <tr id="Tr2" runat="server" style="display: none;">
                    <%--   <asp:BoundField DataField="Colname" HeaderText=""  >  <HeaderStyle Width="200px" />  <%-- <ItemStyle Width="500px"></ItemStyle>    </asp:BoundField>  --%>
                </tr>
                <%--   <asp:BoundField DataField="Colname" HeaderText=""  >  <HeaderStyle Width="200px" />  <%-- <ItemStyle Width="500px"></ItemStyle>    </asp:BoundField>  --%>
                </tr>
                <tr id="trRelationshipFee" runat="server" style="display: none;">
                    <td class="auto-style57">
                        <asp:Label ID="lblRelationshipFee" runat="server" Text="Relationship Fee" Visible="false"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">
                        <asp:TextBox ID="txtRelationshipFee" runat="server" OnKeyPress="return isNumberKey(this, event);" AutoPostBack="True" OnTextChanged="txtRelationshipFee_TextChanged" Visible="false" CssClass="TextAlgRgh"></asp:TextBox>
                        <%--  <tr id="Tr1" runat="server" >
                 <td class="auto-style57">
                        <asp:Label ID="lblbox1" runat="server" Text="Relationship Fees" ></asp:Label></td>
                         <td class="auto-style58">&nbsp;&nbsp;</td>
                 <td class="auto-style59">
                        <asp:TextBox ID="txtBox1" runat="server" 
                            OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7"  CssClass="TextAlgRgh"
                            AutoPostBack="True" ontextchanged="txtBox1_TextChanged"></asp:TextBox>
                       <%-- <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtBox1" Display="Dynamic" ErrorMessage="Enter Fee Amount" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    
                    </td>
                     </tr>--%>
                    </td>

                </tr>



                <tr>
                    <td class="auto-style1">

                        <asp:Label ID="lblCustFee" runat="server"
                            Text="Custom Fee Amount"></asp:Label>
                    </td>
                    <td class="auto-style1"></td>
                    <td class="auto-style1">
                        <asp:TextBox ID="txtCustFeeAmount" runat="server" OnKeyPress="return isNumberKey(this, event);" AutoPostBack="True" OnTextChanged="txtCustFeeAmount_TextChanged" TabIndex="7" CssClass="TextAlgRgh"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server"
                            ControlToValidate="txtCustFeeAmount" Display="Dynamic"
                            ErrorMessage="Enter custom fee" SetFocusOnError="True" ValidationGroup="vgBIE"
                            Visible="false">*</asp:RequiredFieldValidator>
                    </td>
                    <td>

                        <asp:Label ID="lbldiscount" runat="server" Text="Discount" Font-Bold="true"></asp:Label>

                    </td>

                    <td>

                        <asp:TextBox ID="txtDiscount" runat="server" AutoPostBack="True"
                            OnKeyPress="return isNumberKey(this, event);"
                            OnTextChanged="txtDiscount_TextChanged" CssClass="TextAlgRgh"></asp:TextBox>

                    </td>


                    <td>

                        <asp:TextBox ID="txtBillingAUM" runat="server" AutoPostBack="True" CssClass="TextAlgRgh" OnKeyPress="return isNumberKey(this, event);" OnTextChanged="txtBillingAUM_TextChanged" TabIndex="6"></asp:TextBox>
                        <%--  <td class="auto-style57">
                   
                 </td>
                         <td class="auto-style58">&nbsp;</td>
                 <td class="auto-style59">
                        <asp:TextBox ID="txtBox2" runat="server" 
                            OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7"  CssClass="TextAlgRgh"
                            AutoPostBack="True" ontextchanged="txtBox2_TextChanged" ></asp:TextBox>
                       <%-- <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ControlToValidate="txtBox2" Display="Dynamic" ErrorMessage="Enter Fee Amount" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>--%>
                    
                    </td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td></td>
                    <td colspan="3">
                        <asp:Label ID="Label2" runat="server" Text="Leave blank,if using the standard fee amount" Font-Italic="true" ForeColor="Red"></asp:Label>
                    </td>
                    <td>&nbsp;
                         &nbsp;
                         <asp:Label ID="lbldecimal" runat="server" Font-Italic="true" ForeColor="Red"
                             Text="input as Decimal"></asp:Label>

                    </td>
                </tr>

                <tr>
                    <td class="auto-style57">
                        <asp:Label ID="Label4" runat="server" Text="Total Annual Fee Cal"></asp:Label>
                    </td>
                    <td class="auto-style58">&nbsp;</td>
                    <td>
                        <asp:TextBox ID="txtTotalAnnualFee" runat="server"
                            OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7" CssClass="TextAlgRgh">
                        </asp:TextBox>

                    </td>

                    <td>
                        <asp:Label ID="lblFeeRateCalc" runat="server" Text="Fee Rate Calc" Font-Bold="true"></asp:Label></td>
                    <td>


                        <asp:TextBox ID="txtFeeRateCalc" runat="server" AutoPostBack="True" OnTextChanged="txtFeeRateCalc_TextChanged" BackColor="#C7EDF7" CssClass="TextAlgRgh"></asp:TextBox>

                        <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ControlToValidate="txtFeeRateCalc" Display="Dynamic" ErrorMessage="Enter fee rate calc" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>

                </tr>
                <tr>
                    <td class="auto-style57">
                        <asp:Label ID="Label1" runat="server" Text="Quarterly Fee Calc"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59" colspan="3">
                        <asp:TextBox ID="txtQuaterlyFeeCalc" runat="server" OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7" CssClass="TextAlgRgh"></asp:TextBox>
                        &nbsp;&nbsp;   
                        <a id="linkAdjustment" href="#" onclick="showAdjustmentSection();return false;">Enter Qtrly Fee Adjustment</a>
                    </td>

                </tr>
                <tr>
                    <td class="auto-style57">
                        <asp:Label ID="lblFeePerMonth" runat="server" Text="Fees per Month in Qtr"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">
                        <asp:TextBox ID="txtFeesPerMonth" runat="server" OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7" CssClass="TextAlgRgh"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" ControlToValidate="txtFeesPerMonth" Display="Dynamic" ErrorMessage="Enter fees per month in Qtr" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                        <asp:CheckBox ID="chkAccured" runat="server" Text="Accrued(check if this is a Accrued Fee)" />
                    </td>

                </tr>
                <tr id="trAdjustmentAmt" style="display: none">
                    <td class="auto-style57">
                        <asp:Label ID="lblAdjustmentAmt" runat="server" Text="Adjustment Amount"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">
                        <asp:TextBox ID="txtAdjAmt" runat="server" OnKeyPress="return isNumberKey(this, event);" AutoPostBack="True" OnTextChanged="txtAdjAmt_TextChanged" CssClass="TextAlgRgh"></asp:TextBox>
                    </td>


                </tr>
                <tr id="trAdjustmentReason" style="display: none">
                    <td class="auto-style57">
                        <asp:Label ID="lblAdjustmentReason" runat="server" Text="Adjustment Reason"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59" colspan="5">
                        <asp:TextBox ID="txtAdjReason" runat="server" Width="385px" MaxLength="40"></asp:TextBox>
                    </td>
                </tr>
                <tr id="trAdjQtrFee" style="display: none">
                    <td class="auto-style57">
                        <asp:Label ID="lblAdjQtrFee" runat="server" Text="Adjusted Quarterly Fee"></asp:Label></td>
                    <td class="auto-style58">&nbsp;</td>
                    <td class="auto-style59">
                        <asp:TextBox ID="txtAdjQtrFee" runat="server" OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7" OnTextChanged="txtAdjQtrFee_TextChanged" CssClass="TextAlgRgh"></asp:TextBox>
                    </td>


                </tr>
                <tr>
                    <td class="auto-style52">

                        <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="txtBillingAUM" Display="Dynamic" ErrorMessage="Enter Billing AUM" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%>&nbsp;</td>
                    <td class="auto-style53">&nbsp;</td>
                    <td class="auto-style54">
                        <asp:HiddenField ID="Hcheckadjvalue" runat="server"
                            OnValueChanged="Hcheckadjvalue_ValueChanged" />
                        &nbsp;</td>


                </tr>
                <tr>
                    <td class="auto-style46">&nbsp;</td>
                    <td class="auto-style47">&nbsp;</td>
                    <td class="auto-style48" colspan="3">
                        <asp:Button ID="btnSubmit" runat="server" Text="Generate Invoice & PDF" ToolTip="Click to generate the Invoice" ValidationGroup="vgBIE" OnClick="btnSubmit_Click" TabIndex="8" />
                        <asp:HiddenField ID="AnnualFee" runat="server" />
                        <br />
                    </td>

                </tr>
            </table>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <asp:Timer ID="Timer1" runat="server" Interval="3000" OnTick="Timer1_Tick" Enabled="False"></asp:Timer>

            <%-- <tr  id="Tr3" runat="server" >
                 <td class="auto-style57">
                        <asp:Label ID="lblbox3" runat="server" Text="Other Fees" ></asp:Label>

                 </td>
                         <td class="auto-style58">&nbsp;</td>
                 <td class="auto-style59">
                        <asp:TextBox ID="txtBox3" runat="server" 
                            OnKeyPress="return isNumberKey(this, event);" BackColor="#C7EDF7"  CssClass="TextAlgRgh"
                            AutoPostBack="True" ontextchanged="txtBox3_TextChanged" ></asp:TextBox>
                       <%-- <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" ControlToValidate="txtBox3" Display="Dynamic" ErrorMessage="Enter fee amount" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>--%>
        </div>

    </form>
</body>
</html>
