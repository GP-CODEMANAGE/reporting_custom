<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmMailPreference.aspx.cs" Inherits="frmTaskNote" %>

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
                    <td colspan="3">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3" class="Titlebig">Gresham Partners, LLC
                    </td>
                </tr>

                <tr>
                    <td>MAIL PREFERENCE REPORT</td>
                </tr>

                <tr>
                    <td colspan="2">
                        <asp:Label ID="lblError" runat="server" ForeColor="Red" ></asp:Label>
                        <br />
                    </td>
                </tr>

                <tr>
                    <td style="width: 25%; height: 25px;">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="lblAssociate" runat="server" Text="Associate:" Font-Names="Verdana"></asp:Label></td>
                    <td style="height: 25px">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:DropDownList Font-Names="Verdana" ID="ddlAssociate" runat="server" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlAssociate_SelectedIndexChanged">
                        </asp:DropDownList></td>
                    <td style="width: 4px; height: 25px"></td>
                </tr>

                <tr>
                    <td style="width: 25%">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="lblHouseHold" runat="server" Text="Household:" Font-Names="Verdana"></asp:Label></td>
                    <td style="height: 40px">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:ListBox Font-Names="Verdana" ID="lstHouseHold" runat="server" Height="220px"
                            Width="220px" AutoPostBack="True" SelectionMode="Multiple"></asp:ListBox></td>
                    <td style="width: 4px; height: 40px"></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label1" runat="server" Text="Mail Type:" Font-Names="Verdana"></asp:Label></td>
                    <td style="height: 40px">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:ListBox Font-Names="Verdana" ID="lstMailType" runat="server" Height="220px"
                            Width="220px" AutoPostBack="True" SelectionMode="Multiple"></asp:ListBox></td>
                    <td style="width: 4px; height: 40px"></td>
                </tr>

                <tr>
                    <td style="width: 25%">

                    </td>
                    <td style="height: 40px">
                        <asp:Button ID="btnsubmit" runat="server" Text="SUBMIT" OnClick="btnsubmit_Click" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
