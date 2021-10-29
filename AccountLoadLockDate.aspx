<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AccountLoadLockDate.aspx.cs" Inherits="AccountLoadLockDate" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Set Data Load Lock Date on Accounts by Source</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

 <script type="text/javascript" language="javascript">
function ValidateDate()
{
    var DataSource=document.getElementById("lstSource").value;

    var LockDate=document.getElementById("txtLockDate").value;
    var AsofDate=document.getElementById("txtAsofdate").value;
    
    if(DataSource=="")
    {
        alert("Please Select Data Source.");
        return false;
    }
//    if(AsofDate=="")
//    {
//        alert("Please enter Accounts with Positions as of Date.");
//        return false;
//    }
    if(LockDate=="")
    {
        alert("Please enter Data Load Lock Date.");
        return false;
    }
      document.getElementById("trSubmit").style.display = "none";
    /*if(startDate!="")
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
    }*/
}
        function ClearLabel()
        {
            document.getElementById("lblError").innerHTML="";
            document.getElementById("trDownLoad").style.display= "none";
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
                                    <asp:Label ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large" Text="Set Data Load Lock Date on accounts by Source"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 18px" valign="top" colspan="3">
                                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label1" runat="server" Text="Feed:"></asp:Label></td>
                                <td style="height: 40px">
                                    <asp:ListBox ID="lstSource" runat="server" Height="110px" onchange="ClearLabel();" SelectionMode="Multiple">
                                    </asp:ListBox></td>
                                <td style="width: 4px; height: 40px;">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label3" runat="server" Text="Accounts with Positions as of:"></asp:Label></td>
                                <td>
                                    <asp:TextBox ID="txtAsofdate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtAsofdate)">
                                        <img id="imgorgDateRec" alt="" onclick="ClearLabel();" border="0" src="images/calander.png" /></a>
                                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsofdate"
                                        Display="None" ErrorMessage="Please enter As Of Date"></asp:RequiredFieldValidator><asp:CustomValidator
                                            ID="CustomValidator1" runat="server" ControlToValidate="txtAsofdate" ErrorMessage="As of date is not valid"
                                            ClientValidationFunction="ValidateForm" Display="None"> </asp:CustomValidator>--%>
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label2" runat="server" Text="Data Load Lock Date:"></asp:Label></td>
                                <td>
                                    <asp:TextBox ID="txtLockDate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtLockDate)">
                                        <img id="img1" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr id="trSubmit" runat="server">
                                <td>
                                </td>
                                <td valign="top">
                                    <br />
                                    <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClientClick="return ValidateDate();" OnClick="btnSubmit_Click" />
                                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                        ShowSummary="False" />
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr runat="server" id="trDownLoad">
                                <td>
                                </td>
                                <td valign="top">
                                    <asp:LinkButton ID="lnkDownLoad" runat="server" OnClick="lnkDownLoad_Click">Download List of Locked Account</asp:LinkButton></td>
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
