<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HouseholdDashboard.aspx.cs" Inherits="HouseholdDashboard" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <%-- <script>
        function reply_click(obj) {
            var id = obj.id;
            alert("Id is " + id);
           // window.open("http://gp-crm2016:9999/ClientServicesDashboard.aspx?id=" + id, entityType, "menubar=0,toolbar=0,resizable=no,fullscreen=no");
            window.location.href = "http://gp-crm2016:9999/ClientServicesDashboard.aspx?id=" + id;
        }
    </script>--%>
    
    <title>Household Dashboard</title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <img src="images/Gresham_Logo__.jpg" />        <%--<a href="link2.aspx" target="_blank" onclick="javascript:window.open('link2.aspx', 'DetailView', 'location=0, status=0, resizable=1, scrollbars=1, width=800, height=800'); return false;">View Details</a>--%>
    </div>
        <br />
        <br />
        <asp:Label ID="lblError" runat="server" Font-Size="Larger" ForeColor="Red" Text="Label" Visible="False"></asp:Label>
        <br />
        <table>
            <tr>
                <td>
                   &nbsp;&nbsp;&nbsp;<asp:Label ID="lblRoles" Font-Bold="true" runat="server" Text="Role:" Font-Size="Large"></asp:Label>
                </td>
                <td>&nbsp;&nbsp;&nbsp;<asp:DropDownList ID="ddlRoles" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlRoles_SelectedIndexChanged" Font-Size="Large"></asp:DropDownList></td>
                <td>
                   &nbsp;&nbsp;&nbsp;<asp:Label ID="lblAdvisor" Font-Bold="true" runat="server" Text="Household Name:  " Font-Size="Large"></asp:Label>
                </td>
                 <td>
                   &nbsp;&nbsp;&nbsp; <asp:DropDownList ID="ddlAdvisor" runat="server" AutoPostBack="true" Font-Size="Large" Visible="true"  >
                     

                    </asp:DropDownList>

                </td>
                 <td> &nbsp;&nbsp;&nbsp;<asp:Button ID="btnReset" runat="server" Text="Reset" OnClick="btnReset_Click" /> </td>
                </tr>
            </table>
         <br />
        

            <asp:Label ID="Label1" runat="server" Text="Label" Visible="False" Font-Size="Large" ForeColor="Red"></asp:Label>
        
    </form>
</body>
</html>
