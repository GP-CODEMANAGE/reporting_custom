<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PerformanceAddUpdateTool_PopUp.aspx.cs"
    Inherits="PerformanceAddUpdateTool_PopUp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <base target="_self" />
    <%--<META HTTP-EQUIV='Pragma' CONTENT='no-cache'">
<META HTTP-EQUIV="Expires" CONTENT="-1">--%>
    <title>Performance Add Tool</title>
    <style type="text/css"> 
    .CellTopBorder
    {
	    border-top-color:Gray; border-top:solid; border-top-width:thick;
	
    }
   
    .displayNone
    {
        display:none;
    }
    
    .CellTitle
    {
       border-bottom: black 1px solid; 
    }
    
    .CellHeader
    {
       border-bottom: black 1px solid; 
       border-left: black 1px solid; 
       border-right: black 1px solid; 
      
    }
    
    .CellTotLeft
    {
        border-bottom: black 1px solid; 
        border-left: black 1px solid;
    }
    
    .CellTotRight
    {
        border-bottom: black 1px solid; 
        border-left: black 1px solid;
        border-right: black 1px solid;
    }
    
    
    .CellTop
    {
        border-top: black 1px solid; 
        border-bottom: black 1px solid; 
        border-left: black 1px solid;
        border-right: black 1px solid;
    }
    
    </style>

    <script type="text/javascript" language="javascript">
    
    function ReturnToParent(value)
    {
        window.returnValue = value;
        window.close();
    }
    
    // For Auto refreshing the grid values
     function Refressh()
     {
     //debugger;
        if (event.keyCode == 13)
        {
          __doPostBack("txtPerformance", "TextChanged");
           return false;
        }
         
     }
    
    function validateCAUpdateValue(grp1)
    {
       var validated = Page_ClientValidate(grp1);
       var frm = document.forms[0];    
       
       var performance
        if (validated)
        {
            document.getElementById("trButton").style.display="none";
        }
        else
        {
            alert('Please enter only numeric values in Performance Value');
            return false;
        }
       
    }
    
    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        &nbsp;<asp:Label ID="lblMessage" runat="server" Font-Bold="True" ForeColor="Red"
                            Visible="False"></asp:Label>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top">
                        <div id="dvGrid" runat="server">
                            <table border="1" style="">
                                <tr>
                                    <td align="center" style="font-weight: bold;">
                                        <asp:Label ID="lblPerfType" runat="server" Text=""></asp:Label></td>
                                    <td align="center" style="font-weight: bold;">
                                        Performance As Of Date
                                    </td>
                                    <td align="center" style="font-weight: bold;">
                                        Performance
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="lblName" runat="server" Text=""></asp:Label>
                                    </td>
                                    <td>
                                        <asp:Label ID="lblDate" runat="server" Text=""></asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtPerformance" runat="server"></asp:TextBox>
                                        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="txtPerformance"
                                            Display="Dynamic" ErrorMessage="Please enter numeric values only" ValidationExpression="^-?\d*(\.\d+)?$"
                                            ValidationGroup="grp1">*</asp:RegularExpressionValidator>
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr id="trButton" runat="server">
                    <td align="right" colspan="2" style="height: 26px">
                        &nbsp;<asp:Button ID="btnSubmit" runat="server" OnClick="btnSubmit_Click" Text="Submit"
                            OnClientClick="return validateCAUpdateValue('grp1');" />&nbsp;
                        <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClientClick="return ReturnToParent(null);return false;" />&nbsp;
                    </td>
                </tr>
                <tr>
                    <td style="height: 21px">
                    </td>
                    <td style="height: 21px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
