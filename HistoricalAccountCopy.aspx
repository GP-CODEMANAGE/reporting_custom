<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HistoricalAccountCopy.aspx.cs"
    Inherits="HistoricalAccountCopy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<link href="common/Calendar.css" rel="stylesheet" type="text/css" />

<script src="common/Calendar.js" type="text/javascript"></script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Historical Account Copy</title>

    <script language="javascript" type="text/javascript">
    
    function Validate()
    {
        var HH =document.getElementById("<%=ddlHouseHold.ClientID %>").value;
        var AlltrPstn =document.getElementById("<%=chkCopyAllTranPostn.ClientID %>").checked;
        if(HH=="" || HH=="0")
        {
            alert("Please select household.");
            return false;
        }
        
        if(AlltrPstn==false)
        {
            var starDt= document.getElementById("<%=txtStartDate.ClientID %>").value;
             var EndDt= document.getElementById("<%=txtEndDate.ClientID %>").value;
             
             if(starDt=="")
             {
                 alert("Please provide start date.");
                 return false;
             }
             if(EndDt=="")
             {
                 alert("Please provide end date.");
                 return false;
             }
             
            if (starDt != "" && EndDt != "")
            {
                var dt1 = parseInt(starDt.substring(3, 5), 10);
                var mon1 = parseInt(starDt.substring(0, 2), 10);
                var yr1 = parseInt(starDt.substring(6, 10), 10);
                var dt2 = parseInt(EndDt.substring(3, 5), 10);
                var mon2 = parseInt(EndDt.substring(0, 2), 10);
                var yr2 = parseInt(EndDt.substring(6, 10), 10);
                var date1 = new Date(yr1, mon1, dt1);
                var date2 = new Date(yr2, mon2, dt2);
                
                if (date2 < date1) {
                    alert("End Date cannot be greater than Start Date");
                    return false;
                }
            }
        }
        
        document.getElementById("trSubmit").style.display = "none";
        document.getElementById("trLoader").innerHTML = "<span style='font-weight:bold;color:Red;'>Loading. Please wait.</span><br /><img src='http://mem02/ISV/AdventReport/Images/ajax-loader.gif' />";
    }
    
    function ClearLabel()
    {
        document.getElementById("lblError").innerHTML="";
        document.getElementById("lnkTransactionErrorFile").style.display = "none";
        document.getElementById("lnkPositionErrorFile").style.display = "none";
    }
    
    function OnHHChange()
    {
        document.getElementById("<%=txtStartDate.ClientID %>").value = "";
        document.getElementById("<%=txtEndDate.ClientID %>").value = "";      
        document.getElementById("<%=chkCopyAllTranPostn.ClientID %>").checked = true;
        ClearLabel();
    }
    
    function OnDateSelection()
    {
        document.getElementById("<%=chkCopyAllTranPostn.ClientID %>").checked = false;
        ClearLabel();
    }
    
    function OnCheckboxClick()
    {
        var AlltrPstn =document.getElementById("<%=chkCopyAllTranPostn.ClientID %>").checked;
          
        if(AlltrPstn==true)
        {
            document.getElementById("<%=txtStartDate.ClientID %>").value = "";
            document.getElementById("<%=txtEndDate.ClientID %>").value = "";
        }
        ClearLabel();
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
                        <strong>Historical Account Copy</strong>
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
                        <asp:DropDownList ID="ddlHouseHold" onchange="OnHHChange();" runat="server">
                        </asp:DropDownList></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        Start Date</td>
                    <td>
                        <asp:TextBox ID="txtStartDate" runat="server" onkeyup="OnDateSelection();"></asp:TextBox>
                        <a onclick="showCalendarControl(txtStartDate)">
                            <img id="img1" alt="" border="0" onclick="OnDateSelection();" src="images/calander.png" /></a>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        End Date</td>
                    <td>
                        <asp:TextBox ID="txtEndDate" runat="server" onkeyup="OnDateSelection();"></asp:TextBox>
                        <a onclick="showCalendarControl(txtEndDate)">
                            <img id="img2" alt="" border="0" onclick="OnDateSelection();" src="images/calander.png" /></a>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        Copy All Positions and Transactions</td>
                    <td>
                        <asp:CheckBox ID="chkCopyAllTranPostn" onclick="OnCheckboxClick();" runat="server" /></td>
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
                    <td style="height: 20px">
                        <asp:LinkButton ID="lnkTransactionErrorFile" runat="server" OnClick="lnkTransactionErrorFile_Click">Download Historical Acct Load Summary File</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:LinkButton ID="lnkPositionErrorFile" runat="server" OnClick="lnkPositionErrorFile_Click">Download Position Exception File</asp:LinkButton>
                    </td>
                    <td>
                    </td>
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
