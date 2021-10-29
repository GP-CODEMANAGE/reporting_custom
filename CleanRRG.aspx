<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CleanRRG.aspx.cs" Inherits="CleanRRG" %>

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
                    <td class="Titlebig">RRG Cleanup
                    </td>
                </tr>
                    <tr>

                    <td style="height: 18px" valign="top" colspan="2">
                        <asp:Label ID="lblMessage" runat="server" ForeColor="Red" Visible="False"></asp:Label><br />

                        <asp:Label ID="lblError" runat="server" Font-Bold="False" ForeColor="Red" Text="lblError" Visible="False"></asp:Label>
                    </td>

                </tr> 
                  <tr>

                    <td style="width: 25%">
                        <asp:Label ID="Label1" runat="server" Text="Household:" Font-Names="Verdana"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlHH" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHH_SelectedIndexChanged"   >
                        </asp:DropDownList></td>
                   <%-- <td style="width: 4px; height: 40px"></td>--%>

                </tr>
                  <tr>

                    <td style="width: 25%">
                        <asp:Label ID="lblHouseHold" runat="server" Text="ReportRollupGroup:" Font-Names="Verdana"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlRRG" runat="server" AutoPostBack="true" OnSelectedIndexChanged="RRG_SelectedIndexChanged"  >
                        </asp:DropDownList></td>
                   <%-- <td style="width: 4px; height: 40px"></td>--%>

                </tr>
                    <tr>
                                <td style="width: 25%; font-family: Verdana">Legal Entity</td>
                                <td style="height: 40px; width: 720px;">
                                    <asp:ListBox ID="lstLegalEntity" runat="server" Height="220px" Width="220px" AutoPostBack="True"
                                         SelectionMode="Multiple"
                                        Font-Names="Verdana"></asp:ListBox></td>
                            </tr>
                     <tr>
                    <td></td>
                    <td>
                        <asp:Button ID="btnSumbit" Font-Names="Verdana" runat="server" 
                           Text="Create ReportRollupGroup" Enabled="True" OnClick="btnSumbit_Click" />
                    
                        <asp:Button ID="btnLookthroughAccount" Font-Names="Verdana" runat="server" 
                           Text="Create LookthroughAccount" OnClick="btnLookthroughAccount_Click" />
                    </td>
                </tr>
                  </table>
        </div>
    </form>
</body>
</html>
