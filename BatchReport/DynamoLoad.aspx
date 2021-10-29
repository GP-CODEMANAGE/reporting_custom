<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DynamoLoad.aspx.cs" Inherits="DynamoLoad" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Data Load : Dynamo To CRM</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <script type="text/javascript" language="javascript">


        function hideRow() {
            debugger;


            var hidevar = document.getElementById("uplhide");

            if (hidevar.value != 1) {
                var chkTransaction = document.getElementById("UplTransaction").files.length;
                var chkSummaryValuation = document.getElementById("UplSummaryValuation").files.length;
                var chkPosition = document.getElementById("UPLPosition").files.length;
                var chkPerformance = document.getElementById("UPLPerf").files.length;

                //  alert(chkTransaction);
                if (chkTransaction == 0 && chkSummaryValuation == 0 && chkPosition == 0 && chkPerformance == 0) {
                    alert("Please UpLoad atleast 1 file.");
                    return false;
                }

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
            <asp:MultiView ID="mvLoad" runat="server">
                <asp:View ID="vRunLoad" runat="server">
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
                                <asp:Label ID="lbltrans" runat="server" ForeColor="Red"></asp:Label>

                            </td>

                            <%-- <td style="height: 18px" valign="top">
                                <asp:Label ID="Label1" runat="server" ForeColor="Red"></asp:Label>
                               
                            </td>--%>
                        </tr>

                        <tr>
                            <td style="height: 18px" valign="top" colspan="3">
                                <asp:Label ID="lblcommit" runat="server" ForeColor="Red"></asp:Label>

                            </td>
                        </tr>

                        <tr>
                            <td style="height: 18px" valign="top" colspan="3">
                                <asp:Label ID="lblsummary" runat="server" ForeColor="Red"></asp:Label>

                            </td>
                        </tr>

                        <tr>
                            <td style="height: 18px" valign="top" colspan="3">
                                <asp:Label ID="lblperf" runat="server" ForeColor="Red"></asp:Label>

                            </td>
                        </tr>

                         <tr>
                            <td style="height: 18px" valign="top" colspan="3">
                                <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>

                            </td>
                        </tr>

                        <tr>
                            <td class="Titlebig" colspan="3">Data Load From Dynamo to CRM</td>
                        </tr>

                        <tr>
                            <td style="width: 182px">Transaction Load</td>
                            <td>

                                <asp:FileUpload ID="UplTransaction" runat="server" Width="483px" />

                                <asp:LinkButton ID="linktransaction" runat="server" Text="Sample File Transaction" OnClick="linktransaction_Click"></asp:LinkButton>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td style="width: 182px">Summary Valuation Load</td>
                            <td>

                                <asp:FileUpload ID="UplSummaryValuation" runat="server" Width="483px" />

                                <asp:LinkButton ID="linksumaary" runat="server" Text="Sample File Summary valuation" OnClick="linksumaary_Click"></asp:LinkButton>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td style="width: 182px">Position Load</td>
                            <td>

                                <asp:FileUpload ID="UPLPosition" runat="server" Width="483px" />

                                <asp:LinkButton ID="linkPostion" runat="server" Text="Sample File Commitment" OnClick="linkPostion_Click"></asp:LinkButton>
                            </td>
                            <td></td>
                        </tr>

                        <%--<tr>
                            <td style="width: 182px">Performance Load</td>
                            <td>

                                <asp:FileUpload ID="FileUpload1" runat="server" Width="486px" />
                            </td>
                            <td></td>
                        </tr>--%>


                        <tr>
                            <td style="width: 182px">Performance Load</td>
                            <td>

                                <asp:FileUpload ID="UPLPerf" runat="server" Width="483px" />

                                <asp:LinkButton ID="linkPerformance" runat="server" Text="Sample File Performance" OnClick="linkPerformance_Click"></asp:LinkButton>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td style="height: 40px"></td>
                            <td style="height: 40px" id="Td1" runat="server">
                                <%--<asp:Label ID="lblMessage" Text="Please wait .... Data Load is in process" runat="server"></asp:Label>--%>
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

                        <tr >
                            <td style="width: 182px"></td>
                            <td>
                                <%--<asp:Button ID="btnLoadRun" runat="server" Text="Run Load" OnClientClick="return hideRow();"
                                    OnClick="btnLoadRun_Click" />--%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                               <%-- <asp:Button ID="Button1" runat="server" Text="Lock" OnClientClick="return hideRow();"
                                    OnClick="btnLock_Click" />--%>


                            </td>

                        </tr>




                    </table>
                    <table>

                        <tr id="trSubmit" runat="server">
                             <td style="width: 182px"></td>

                            <td style="width: 182px;">
                                <asp:Button ID="btnLoadRun" runat="server" Text="Run Load" OnClientClick="return hideRow();"
                                    OnClick="btnLoadRun_Click" />

                               <%-- <asp:Button ID="Button1" runat="server" Text="Lock" OnClientClick="return hideRow();"
                                    OnClick="btnLock_Click" />--%>


                            </td>


                            <td style="width: 182px; text-align: center">
                                <asp:Button ID="btnLock" runat="server" Text="Lock" OnClientClick="return hideRow();"
                                    OnClick="btnLock_Click" /></td>


                            <td>
                                <input id="ConfirmLock" runat="server" value="0" type="hidden" />
                            </td>


                        </tr>


                        <tr runat="server" id="trDownLoad">
                            <td style="width: 182px"></td>
                            <td>
                                <asp:LinkButton ID="lnkDownLoad" runat="server" OnClick="lnkDownLoad_Click" Visible="false">Download Missing Account/Security</asp:LinkButton>
                                <asp:Button ID="btnContinue" runat="server" Text="Continue" OnClientClick="return hideRow();"
                                    OnClick="btnContinue_Click" />
                                <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" /></td>
                            <td></td>
                        </tr>

                        <tr id="Tr1" runat="Server">
                            <td style="height: 40px"></td>
                            <td style="height: 40px" id="trLoader" runat="server">
                                <%--<asp:Label ID="lblMessage" Text="Please wait .... Data Load is in process" runat="server"></asp:Label>--%>
                            </td>
                            <td style="height: 40px"></td>
                        </tr>

                        <asp:HiddenField ID="uplhide" runat="server" />
                    </table>
                </asp:View>

            </asp:MultiView>
        </div>
    </form>
</body>
</html>
