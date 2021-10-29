<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AssetsPerformanceSummary_New.aspx.cs" Inherits="AssetsPerformanceSummary_New" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Gresham Advised Assets Performance Summary Report GA 2.0</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

 <script type="text/javascript" language="javascript">
function ValidateDate()
{
    var Household=document.getElementById("ddlHousehold").value;
    var Group=document.getElementById("ddlGroup").value;

    var startDate=document.getElementById("txtStartDate").value;
    var AsofDate=document.getElementById("txtAsofdate").value;
    
    if(Household=="")
    {
        alert("Please Select HouseHold.");
        return false;
    }
    if(Group=="")
    {
        alert("Please Select Group.");
        return false;
    }
    if(AsofDate=="")
    {
        alert("Please enter As Of Date.");
        return false;
    }
    if(startDate!="")
    {
        startDate=startDate.split('/');
        AsofDate=AsofDate.split('/');

        var SDate=new Date();
        SDate.setFullYear(startDate[2],startDate[0],startDate[1]);

        var AFDate=new Date();
        AFDate.setFullYear(AsofDate[2],AsofDate[0],AsofDate[1]);
        
            if (SDate>=AFDate)
              {
               alert("Please pick an as of date that is after your selected start date.");
               return false;
              }
    }
}
    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 100%">
                <tr>
                    <td>
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
                                <td colspan="3" align="left">
                                    <asp:Label ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large" Text="Gresham Advised Assets Performance Summary Report GA 2.0"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 18px" valign="top" colspan="3">
                                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                            <tr>
                                <td style="width: 20%; height: 26px;">
                                    <asp:Label ID="lblHouseHold" runat="server" Text="HouseHold"></asp:Label></td>
                                <td style="width: 80%; height: 26px;">
                                    <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                                    </asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlHousehold"
                                        Display="None" ErrorMessage="Please Select HouseHold"></asp:RequiredFieldValidator>
                                </td>
                                <td style="width: 4px; height: 26px;">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label11" runat="server" Text="Group:"></asp:Label></td>
                                <td style="height: 40px">
                                    <asp:DropDownList ID="ddlGroup" runat="server">
                                    </asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlGroup"
                                        Display="None" ErrorMessage="Please Select Group"></asp:RequiredFieldValidator>
                                </td>
                                <td style="width: 4px; height: 40px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label1" runat="server" Text="Asset Class:"></asp:Label></td>
                                <td style="height: 40px">
                                    <asp:ListBox ID="lstAssetClass" runat="server" Height="170px" SelectionMode="Multiple">
                                    </asp:ListBox></td>
                                <td style="width: 4px; height: 40px;">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label3" runat="server" Text="As Of Date:"></asp:Label></td>
                                <td>
                                    <asp:TextBox ID="txtAsofdate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtAsofdate)">
                                        <img id="imgorgDateRec" alt="" border="0" src="images/calander.png" /></a>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsofdate"
                                        Display="None" ErrorMessage="Please enter As Of Date"></asp:RequiredFieldValidator><asp:CustomValidator
                                            ID="CustomValidator1" runat="server" ControlToValidate="txtAsofdate" ErrorMessage="As of date is not valid"
                                            ClientValidationFunction="ValidateForm" Display="None"> </asp:CustomValidator>
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label2" runat="server" Text="Start Date:"></asp:Label></td>
                                <td>
                                    <asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtStartDate)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" /></a>
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td>
                                </td>
                                <td valign="top">
                                    <br />
                                    <asp:Button ID="Button1" runat="server" Text="Generate Report" OnClientClick="return ValidateDate();" OnClick="Button1_Click" />
                                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                        ShowSummary="False" />
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
