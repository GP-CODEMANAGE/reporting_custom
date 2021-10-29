<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HistoricalAccountRemoval.aspx.cs"
    Inherits="HistoricalAccountRemoval" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Historical Account Removal</title>

    <script language="javascript" type="text/javascript">
    
    function Validate()
    {
        var HH =document.getElementById("<%=ddlHouseHold.ClientID %>").value;
        if(HH=="" || HH=="0")
        {
            alert("Please select household.");
            return false;
        }
        
        var result=confirm("You are deleting ALL positions, transactions and accounts in the Historical Load Accounts legal entity for the Household selected.  Do you want to continue?");
        if(result)
        {
            document.getElementById("trSubmit").style.display = "none";
            document.getElementById("trLoader").innerHTML = "<span style='font-weight:bold;color:Red;'>Loading. Please wait.</span><br /><img src='http://mem02/ISV/AdventReport/Images/ajax-loader.gif' />";
            return true;
        }
        else
        {
            document.getElementById("<%=ddlHouseHold.ClientID %>").value = "0";
            return false;
        }    
    }
        function ClearLabel()
    {
        document.getElementById("lblError").innerHTML="";
    }
    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table width="100%">
                <tr>
                    <td colspan="3">
                        <img alt="" src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3" class="Titlebig">
                        Gresham Partners, LLC
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <strong>Historical Account Removal</strong>
                    </td>
                </tr>
                <tr>
                    <td valign="top" colspan="3">
                        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 20%;">
                        <asp:Label ID="lblHouseHold" runat="server" Text="HouseHold"></asp:Label></td>
                    <td style="width: 30%;">
                        <asp:DropDownList ID="ddlHouseHold" onchange="ClearLabel();" runat="server">
                        </asp:DropDownList></td>
                    <td>
                    </td>
                </tr>
                <tr runat="server" id="tr2">
                    <td style="height:20px;">
                        &nbsp;</td>
                    <td>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr runat="server" id="trSubmit">
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="btnSubmit" runat="server" OnClientClick="return Validate();" Text="Submit"
                            OnClick="btnSubmit_Click" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr runat="server" id="tr1">
                    <td>
                        &nbsp;</td>
                    <td style="height: 40px" id="trLoader" runat="server">
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td valign="top" style="width: 185px">
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                            ShowSummary="False" />
                    </td>
                    <td style="width: 4px">
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
