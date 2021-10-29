<%@ Page Language="C#" AutoEventWireup="true" CodeFile="RRGExcel.aspx.cs" Inherits="DumpExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>

            <table>
                <tr>
                    <td>
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td class="Titlebig">Gresham Partners, LLC
                    </td>
                </tr>
                <tr>
                    <td class="Titlebig" colspan="3">
                        <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Large"
                            Text="Excel Report" Width="260px"></asp:Label>

                    </td>
                </tr>
                <tr>

                    <td style="height: 18px" valign="top" colspan="2">
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label><br />

                        <asp:Label ID="lblError" runat="server" Font-Bold="False" ForeColor="Red" Text="lblError" Visible="False"></asp:Label>
                    </td>

                </tr>
                <tr>

                    <td style="width: 25%">
                        <asp:Label ID="lblHouseHold" runat="server" Text="Household:" Font-Names="Verdana"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlHouseHold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHouseHold_SelectedIndexChanged" >
                        </asp:DropDownList></td>
                   <%-- <td style="width: 4px; height: 40px"></td>--%>

                </tr>
                <tr>

                    <td style="width: 25%">
                        <asp:Label ID="lblGAgroup" runat="server" Text="GA Group" Font-Names="Verdana"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlGAGroup" runat="server" AutoPostBack="false" >
                        </asp:DropDownList></td>
                    <%--<td style="width: 4px; height: 40px"></td>--%>

                </tr>
                <tr>

                    <td style="width: 25%">
                        <asp:Label ID="Label2" runat="server" Text="TIA Group" Font-Names="Verdana"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlTIAGroup" runat="server" AutoPostBack="false">
                        </asp:DropDownList>

                    </td>
                    <%--<td style="width: 4px; height: 40px"></td>--%>

                </tr>
                <tr>
                    <td></td>
                    <td>
                        <asp:Button ID="btnSumbitTop" Font-Names="Verdana" runat="server" 
                            OnClick="btnSubmit_Click" Text="Submit" Enabled="False" />
                    </td>

                </tr>
            </table>

        </div>
    </form>
</body>
</html>
