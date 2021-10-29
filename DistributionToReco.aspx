<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DistributionToReco.aspx.cs"
    Inherits="DistributionToReco" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Upload Cap Calls/ Distribution into Recommendation</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <script src="Scripts/jquery-1.6.4.min.js"></script>

    <style type="text/css">
        .CellTopBorder
        {
            border-top-color: Gray;
            border-top: solid;
            border-top-width: thick;
        }
    </style>
    <script type="text/javascript">

        function txtDateClear() {
            
            document.getElementById('lblError').innerText = '';
             document.getElementById('lblError1').innerText = '';
           
            return true;
        }
    </script>

    <script type = "text/javascript">
        function Confirm() {
            var confirm_value = document.createElement("INPUT");
            confirm_value.type = "hidden";
            confirm_value.name = "confirm_value";
            if (confirm("Do you want to save data?")) {
                confirm_value.value = "Yes";
            } else {
                confirm_value.value = "No";
            }
            document.forms[0].appendChild(confirm_value);
        }
    </script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width:auto">
            <tr>
                <td colspan="3">
                    <img src="images/Gresham_Logo__.jpg" />
                </td>
            </tr>
            <tr>
                <td colspan="3" >
                    Gresham Partners, LLC
                </td>
            </tr>
            <tr>
                <td  colspan="3">
                    Upload Cap Calls/ Distribution into Recommendation
                </td>
            </tr>
            <tr>
                <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                </td>

               
               <%-- <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError2" runat="server" ForeColor="Red"></asp:Label>
                </td>
                <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError3" runat="server" ForeColor="Red"></asp:Label>
                </td>
                <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError4" runat="server" ForeColor="Red"></asp:Label>
                </td>--%>
            </tr>

            <tr>
                  <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError1" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>

            <tr>
                <td style="text-align: center">
                    <strong>Capital Call/Distribution Date </strong>
                </td>
                <td>
                    <asp:TextBox ID="txtDate" runat="server" Width="164px"  onclick="txtDateClear();" ></asp:TextBox>
                    <a onclick="showCalendarControl(txtDate)">
                        <img id="img1"  border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator
                            ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtDate" Display="Dynamic"
                            ErrorMessage="Please Enter Date" ValidationGroup="vgTOI"></asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td style="text-align: center">
                    <strong>Select File</strong>
                </td>
                <td>
                    <asp:FileUpload ID="fuDist" runat="server" Width="260px" OnClientClick="txtDateClear();"   />
                    <%--<asp:RegularExpressionValidator ID="RegularExpressionValidator1" ValidationExpression="([a-zA-Z0-9\s_\\.\-:])+(.zip)$"
                        ControlToValidate="fuDist" runat="server" ForeColor="Red" ErrorMessage="Please Select zip File"
                        Display="Dynamic" ValidationGroup="vgTOI" />--%>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="fuDist" Display="Dynamic"
                            ErrorMessage="Please Select File" ValidationGroup="vgTOI"></asp:RequiredFieldValidator>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td style="text-align: center">
                    <strong>Select File Type</strong>
                </td>
                <td>
                    <asp:Panel ID="PnlFileType" runat="server" Height="24px">
                        &nbsp;<asp:RadioButton ID="rbCapitalCall" runat="server" Text="Capital Call"     onclick="txtDateClear();"
                            GroupName="TypesOfFiles" Width="112px"  />
                        <asp:RadioButton ID="rbDistribution" runat="server" Text="Distribution "   onclick="txtDateClear();"
                            GroupName="TypesOfFiles" Width="120px" />
                    </asp:Panel>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td style="text-align: center">
                    <asp:Button ID="btnUpload" runat="server" Text="Upload File" Width="114px" ForeColor="Black"
                        BorderColor="Black" OnClick="btnUpload_Click" ValidationGroup="vgTOI" OnClientClick="txtDateClear();" />
                </td>
                <td>
                </td>
            </tr>
         
        </table>

        
        <asp:Button ID="Button2" runat="server" Text="" OnClick="Button2_Click" Visible="false" Width="1px" BackColor="#FFFFFF" BorderStyle="None" />
        <asp:Button ID="Button1" runat="server" Text="" OnClick="Button1_Click1" Visible="false" Width="1px" BackColor="#FFFFFF" BorderStyle="None" />

        <asp:Button ID="Button3" runat="server" Text=""  OnClick="Button3_Click" Visible="false" Width="1px" BackColor="#FFFFFF" BorderStyle="None" />
        <input id="Hidden1" type="hidden" runat="Server" />
        <input id="Hidden2" type="hidden" runat="Server" /> 
         <input id="Hidden3" type="hidden" runat="Server" /> 
    </div>
    </form>
</body>
</html>
