<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmBillingAUM.aspx.cs" Inherits="frmBillingAUM" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Billing: AUM Report</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <style type="text/css">

    input,select
    {
	    font-family:Frutiger 55 Roman;
	    font-size:12pt;
	 
    }
    </style>

    <script type="text/javascript" lang="javascript">
        function ClearValues() {
            document.getElementById("txtAsOfDate").value = "";
        }
    </script>
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
                <td colspan="3" class="Titlebig">
                    Gresham Partners, LLC
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <b>Billing: AUM Report</b>
                </td>
            </tr>
            <tr>
                <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="height: 26px; width: 281px;">
                    AUM
                    As Of &nbsp;Date :
                </td>
                <td style="height: 26px">
                    <asp:TextBox ID="txtAsOfDate" runat="server" Width="119px"></asp:TextBox>
                            <a onclick="showCalendarControl(txtAsOfDate)">
                    <img id="img2" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtAsOfDate"
                        ErrorMessage="Please enter Date"></asp:RequiredFieldValidator>
                </td>
            </tr>
         
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnGenerateReport" runat="server" Text="Generate Report" OnClick="btnGenerateReport_Click" Height="28px" />
                </td>
                <td>
                </td>
            </tr>
    </div>
    </form>
</body></html>

