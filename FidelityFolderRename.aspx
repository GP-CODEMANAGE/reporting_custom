<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FidelityFolderRename.aspx.cs" Inherits="FidelityFolderRename" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <script type="text/javascript">
        function do_totals1() {
            document.all.pleasewaitScreen.style.visibility = "visible";
            window.setTimeout('do_totals2()', 1)
        }
        function do_totals2() {
            calc_totals();
            document.all.pleasewaitScreen.style.visibility = "hidden";
        }

        function hideRow() {
            debugger;
            
            var chkZipFile = document.getElementById("UPLFF").files.length;
            //  alert(chkTransaction);
            if (chkZipFile == 0) {
                alert("Please UpLoad file.");
                return false;
            }
            else {
                return do_totals1();
            }
            function ClearLabel() {
                document.getElementById("lblError").value = "";
            }
        }

    </script>
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
                    <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblTotalFileCount" runat="server" ForeColor="Red"></asp:Label>

                    </td>
                </tr>

                <tr>
                    <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblSucessFileCount" runat="server" ForeColor="Red"></asp:Label>

                    </td>
                </tr>

                <tr>
                    <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblfailFileCount" runat="server" ForeColor="Red"></asp:Label>

                    </td>
                </tr>
                 <tr>
                    <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>

                    </td>
                </tr>

                <%-- <tr>
                    <td class="Titlebig" colspan="3">Data Load From Dynamo to CRM</td>
                </tr>--%>



                <%--<tr>
                            <td style="width: 182px">Performance Load</td>
                            <td>

                                <asp:FileUpload ID="FileUpload1" runat="server" Width="486px" />
                            </td>
                            <td></td>
                        </tr>--%>


                <tr>
                    <td style="width: 182px">Fidelity Folder Upload</td>
                    <td>

                        <asp:FileUpload ID="UPLFF" runat="server" Width="483px" />

                    </td>
                    <td></td>
                </tr>

                <tr>
                    <td style="width: 182px">Save to SarePoint</td>
                    <td>

                        <asp:CheckBox ID="chkSaveSharepoint" runat="server" OnCheckedChanged="chkSaveSharepoint_CheckedChanged" AutoPostBack="true" />

                    </td>
                    <td></td>
                </tr>


                <tr id="tryear" runat="server" visible="false">
                    <td style="width: 182px">Year</td>
                    <td>

                        <asp:TextBox ID="txtyear" runat="server"></asp:TextBox>

                          &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lbldoctype" Text="DOC Type" runat="server"></asp:Label>

                          &nbsp;&nbsp;&nbsp;

                          <asp:DropDownList ID="ddldoctype" runat="server">
                            <asp:ListItem Value="1">1099</asp:ListItem>
                            <asp:ListItem Value="2">5498</asp:ListItem>
                        </asp:DropDownList>

                    </td>
                    <td></td>
                </tr>


             <%--   <tr id="trdoctype" runat="server" visible="false">
                    <td style="width: 182px">Year</td>
                    <td>

                        <asp:DropDownList ID="ddldoctype" runat="server">
                            <asp:ListItem Value="0">select</asp:ListItem>
                            <asp:ListItem Value="1">1099</asp:ListItem>
                            <asp:ListItem Value="2">5498</asp:ListItem>
                        </asp:DropDownList>


                    </td>
                    <td></td>
                </tr>--%>

                <tr>
                    <td style="height: 40px"></td>
                    <td style="height: 40px" id="Td1" runat="server">
                        <asp:LinkButton ID="lnkdownload" runat="server" Visible="false" Text="Download File" OnClick="lnkdownload_Click"></asp:LinkButton>
                    </td>
                    <td style="height: 40px"></td>
                </tr>

                <tr>

                    <%-- <td style="width: 182px">
                                Load Type</td>
                            <td>
                                <asp:CheckBox ID="chkTransaction" Text="Transaction Load" onclick="ClearLabel();" runat="server" Checked="true" />
                                <asp:CheckBox ID="chkSummary" Text="Summary Valuation Load" onclick="ClearLabel();" runat="server" Checked="true" />
                                <asp:CheckBox ID="chkPosition" Text="Position Load" onclick="ClearLabel();" runat="server" Checked="true" /></td>
                            <td>
                            </td>--%>
                </tr>

                <%--<tr>
                            <td style="width: 182px">Position Load</td>
                            <td>

                                <asp:FileUpload ID="FileUpload1" runat="server" Width="486px" />
                            </td>
                            <td></td>
                        </tr>--%>

                <tr>
                    <td style="width: 182px"></td>
                    <td>

                        <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClick="btnSubmit_Click" OnClientClick="return hideRow();" />

                    </td>

                </tr>




            </table>

        </div>


        <div id="pleasewaitScreen" style="z-index: 5; left: 35%; visibility: hidden; position: absolute; top: 40%; width: 256px; height: 191px;">
            <table border="1" style="width: 256px; height: 191px">
                <tr>
                    <td valign="middle" align="center" width="100%" bgcolor="#ffffff" height="100%">
                        <br>
                        <br>
                        <img src="Images/ajax-loader.gif" align="middle"><%-- <FONT face="Lucida Grande, Verdana, Arial, sans-serif" color="#000066" size="5">
           <B> Please wait...</FONT>--%>
                        <br>
                        <br>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
