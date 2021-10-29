<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PerformanceAddUpdateTool.aspx.cs"
    Inherits="PerformanceAddUpdateTool" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <%--<META HTTP-EQUIV='Pragma' CONTENT='no-cache'">
<META HTTP-EQUIV="Expires" CONTENT="-1">--%>
    <title>Performance Add/Update Tool</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <style type="text/css">
        .CellTopBorder {
            border-top-color: Gray;
            border-top: solid;
            border-top-width: thick;
        }

        .displayNone {
            display: none;
        }

        .CellTitle {
            border-bottom: black 1px solid;
        }

        .CellHeader {
            border-bottom: black 1px solid;
            border-left: black 1px solid;
            border-right: black 1px solid;
        }

        .CellTotLeft {
            border-bottom: black 1px solid;
            border-left: black 1px solid;
        }

        .CellTotRight {
            border-bottom: black 1px solid;
            border-left: black 1px solid;
            border-right: black 1px solid;
        }


        .link {
            color: #0000ff;
            font: 7pt verdana normal;
            text-decoration: underline;
            cursor: hand;
        }
    </style>

    <style type="text/css">
        .modalBackground {
            background-color: #333333;
            filter: alpha(opacity=70);
            opacity: 0.7;
            z-index: 9999;
        }

        .modalPopup {
            background-color: #FFFFFF;
            border-width: 1px;
            border-style: solid;
            border-color: #CCCCCC;
            padding: 3px;
            width: auto;
            height: auto;
            margin: 20px;
        }

        .modalPopupSecurity {
            background-color: #FFFFFF;
            border-width: 1px;
            border-style: solid;
            border-color: #CCCCCC;
            padding: 3px;
            width: auto;
            height: 200px;
            margin: 20px;
            /*z-index: 9999;*/
        }
    </style>

    <script type="text/javascript" language="javascript">
        // For Auto refreshing the grid values
        function Refressh() {
            //debugger;
            if (event.keyCode == 13) {
                __doPostBack("txtAsOfDate", "TextChanged");
                return false;
            }

        }
        function getLastDateOfMonth(Year, Month) {
            return (new Date((new Date(Year, Month, 1)) - 1));
        }
        function MonthLastDate(Month, year, dt) {
            var date = getLastDateOfMonth(year, Month);
            var day = date.getDate();
            if (dt == day)
                return true;
            else
                return false;
        }

        function Validate() {
            var AsofDate = document.getElementById("txtAsOfDate").value;
            var Household = document.getElementById("ddlHousehold").value;
            var Associate = document.getElementById("ddlAssociateOps").value;
            var IsOpsTeamMem = document.getElementById("HdIsOpsTeamMember").value;

            //        if(IsOpsTeamMem=="False" && Household=="0" && Associate=="0")
            //        {
            //            alert("Please select household");
            //            return false;
            //        }
            if (IsOpsTeamMem == "False" && Household == "0") {
                alert("Please select household");
                return false;
            }
            if (AsofDate != "") {
                if (!isGoodDate(AsofDate)) {
                    document.getElementById("txtAsOfDate").value = "";
                    alert("Please provide proper M/D/YYYY or M/D/YY or MM/dd/yyyy format date.");
                    return false;
                }
                var dt1 = parseInt(AsofDate.substring(3, 5), 10);
                var mon1 = parseInt(AsofDate.substring(0, 2), 10);
                var yr1 = parseInt(AsofDate.substring(6, 10), 10);

                var IsToLastDate = MonthLastDate(mon1, yr1, dt1);
                if (IsToLastDate == false) {
                    alert("Please select last date of month.");
                    return false;
                }
            }
            else {
                alert("Please provide as of date.");
                return false;
            }
        }
        function validateCAUpdateValue(grp1) {
            if (grp1 == "grp2") {
                Refressh();
            }
            var validated = Page_ClientValidate(grp1);
            var frm = document.forms[0];

            if (validated) {

            }
            else {
                alert('Please enter only numeric values in Performance Value');
                return false;
            }

        }

        function ChangeColor(id) {
            document.getElementById(id).style.backgroundColor = "yellow";
            document.getElementById("lblMessage").innerHTML = "";
        }

        function FreeFormDate(DateVal) {
            //debugger;
            var AsofDate = document.getElementById(DateVal).value;

            if (AsofDate.length >= 6 && AsofDate.indexOf("/") > 0) {
                AsofDate = AsofDate.split('/');

                var mon = AsofDate[0];
                var day = AsofDate[1];
                var year = AsofDate[2];

                if (mon.length < 2)
                    mon = "0" + mon;
                if (day.length < 2)
                    day = "0" + day;
                if (year.length < 3)
                    year = "20" + year;

                document.getElementById(DateVal).value = mon + "/" + day + "/" + year;
            }
            else {
            }

        }

        function isGoodDate(dt) {
            var reGoodDate = /^(?:(0[1-9]|1[012])[\/.](0[1-9]|[12][0-9]|3[01])[\/.](19|20)[0-9]{2})$/;
            return reGoodDate.test(dt);
        }
    </script>

    <script language="javascript" type="text/javascript">

        function OpenChild(UUId, Type, Name, AsOfDate) {
            var querystr = "uuid=" + UUId + "&type=" + Type + "&name=" + encodeURIComponent(Name) + "&asofdate=" + AsOfDate + "";
            var WinSettings = "dialogHeight: 175px; dialogWidth: 675px;  edge: Raised; center: Yes; status: no;";
            //var myObject = window.showModalDialog("PerformanceAddUpdateTool_PopUp.aspx?uuid=" + UUId, myObject, WinSettings);
            var myObject = window.showModalDialog("PerformanceAddUpdateTool_PopUp.aspx?" + querystr, myObject, WinSettings);

            if (myObject != null) {
                __doPostBack('btnRefresh', myObject);
                return false;
            }
            else {
                return false;
            }
        }


    </script>

    <script language="javascript" type="text/javascript">
        function Postback() {
            // debugger

        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <div>
            <table>
                <tr style="height: 50px;">
                    <td colspan="3">
                        <asp:Label ID="lblHeader" runat="server" ForeColor="#00C0C0" Text="PERFORMANCE UPDATE(S)"
                            Font-Bold="True"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="width: 15%;">
                        <b>Associate/OPS</b>
                    </td>
                    <td style="width: 85%;" colspan="2">
                        <asp:DropDownList ID="ddlAssociateOps" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAssociateOps_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td>
                        <b>Household </b>
                    </td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td>
                        <b>As of Date </b>
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="txtAsOfDate" runat="server" onchange="FreeFormDate('txtAsOfDate');"></asp:TextBox><a
                            onclick="showCalendarControl( txtAsOfDate)">
                            <img id="imgorgDateRec" alt="" border="0" src="images/calander.png" /></a>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <div style="font-style: italic;">
                            Please enter performance as a whole number not decimal format (i.e. 1.1% should
                            be input as 1.10 not .011)
                        </div>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td colspan="2">
                        <asp:CheckBox ID="chkUnsupressAll" runat="server" Text="Unsuppress All" /></td>
                </tr>
                <tr>
                    <td></td>
                    <td colspan="2" style="height: 40px">
                        <asp:Button ID="btnLoadData" runat="server" Text="Load Data" OnClientClick="return Validate();"
                            OnClick="btnLoadData_Click" /></td>
                </tr>
                <tr>
                    <td align="left" colspan="3" style="height: 22px">
                        <asp:Label ID="lblMessage" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" colspan="3" style="height: 22px;">
                        <asp:Button ID="btnCanceltop" runat="server" Text="Cancel" OnClick="btnCanceltop_Click" />&nbsp;
                        <asp:Button ID="btnSumbitTop" runat="server" OnClick="btnSubmit_Click" Text="Submit"
                            OnClientClick="return validateCAUpdateValue('grp1');" />&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3">
                        <%-- <asp:UpdatePanel ID="gridview" runat="server" UpdateMode="Always">
                            <ContentTemplate>--%>

                        <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="False" TabIndex="1"
                            ToolTip="Performance Add/Update Tool" ShowHeader="False" GridLines="None" OnRowDataBound="gvList_RowDataBound" OnRowCommand="gvList_RowCommand"
                            Font-Names="Verdana" Font-Size="X-Small" CellPadding="2" Width="100%" EnableViewState="true">
                            <Columns>
                                <asp:BoundField DataField="Name" HeaderText="Name" HtmlEncode="false"></asp:BoundField>
                                <asp:BoundField DataField="Perf Last Upd On" DataFormatString="{0:MM/dd/yyyy}" HeaderText="Performance Last Update On" />
                                <asp:BoundField DataField="Perf Last Upd By" HeaderText="Performance Last Updated By" />
                                <asp:TemplateField HeaderText="Commitment">
                                    <ItemTemplate>
                                        <%--<a runat="server" id="lnkAdd2" class="link">Add</a>--%>
                                        <asp:LinkButton ID="lnkAdd2" CssClass="link" runat="server" Text="Add" CommandName="linkButton2" CommandArgument="<%# Container.DataItemIndex %>"></asp:LinkButton>
                                        <asp:TextBox runat="server" ID="txtPerformance2" Width="82px" Font-Names="Verdana"
                                            Font-Size="X-Small" /><asp:RegularExpressionValidator ID="RegularExpressionValidator2"
                                                runat="server" ControlToValidate="txtPerformance2" Display="Dynamic" ErrorMessage="Please enter numeric values only"
                                                ValidationExpression="^-?\d*(\.\d+)?$" ValidationGroup="grp1">*</asp:RegularExpressionValidator>
                                        <asp:TextBox runat="server" ID="txtPerfHidden2" Width="22px" CssClass="displayNone" />
                                        <asp:Label ID="lblLockDate2" runat="server" Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        11/30/2012 Performance
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Commitment">
                                    <ItemTemplate>
                                        <%-- <a runat="server" id="lnkAdd1" class="link">Add</a>--%>
                                        <asp:LinkButton ID="lnkAdd1" CssClass="link" runat="server" Text="Add" CommandName="linkButton1" CommandArgument="<%# Container.DataItemIndex %>"></asp:LinkButton>
                                        <asp:TextBox runat="server" ID="txtPerformance1" Width="82px" Font-Names="Verdana"
                                            Font-Size="X-Small" /><asp:RegularExpressionValidator ID="RegularExpressionValidator1"
                                                runat="server" ControlToValidate="txtPerformance1" Display="Dynamic" ErrorMessage="Please enter numeric values only"
                                                ValidationExpression="^-?\d*(\.\d+)?$" ValidationGroup="grp1">*</asp:RegularExpressionValidator>
                                        <asp:TextBox runat="server" ID="txtPerfHidden1" Width="22px" CssClass="displayNone" />
                                        <asp:Label ID="lblLockDate1" runat="server" Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        12/31/2012 Performance
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="CurrAsofDate" HeaderText="CurrAsofDate" Visible="False" />
                                <asp:BoundField DataField="_CurrAsofDateUUID" HeaderText="_CurrAsofDateUUID" Visible="False" />
                                <asp:BoundField DataField="PrevAsOfDate" HeaderText="PrevAsOfDate" Visible="False" />
                                <asp:BoundField DataField="_PrevAsOfDateUUID" HeaderText="_PrevAsOfDateUUID" Visible="False" />
                                <asp:BoundField DataField="_HeaderFlg" HeaderText="_HeaderFlg" Visible="False" />
                                <asp:BoundField DataField="_UUID" HeaderText="_UUID" Visible="False" />
                                <asp:BoundField DataField="Name" HeaderText="Name" Visible="False" />
                                <asp:BoundField DataField="_CurrAsofDate" HeaderText="_CurrAsofDate" Visible="False" />
                                <asp:BoundField DataField="_PrevAsOfDate" HeaderText="_PrevAsOfDate" Visible="False" />
                                <asp:BoundField DataField="_Type" HeaderText="_Type" Visible="False" />
                                <asp:BoundField DataField="_CurrentFlg" HeaderText="_CurrentFlg" Visible="False" />
                                <asp:BoundField DataField="_PreviousFlg" HeaderText="_PreviousFlg" Visible="False" />
                                <asp:BoundField DataField="_PerfLockDate" HeaderText="_PerfLockDate" Visible="False" />
                            </Columns>
                            <%--<HeaderStyle Height="10px" BackColor="#BFDBFF" Font-Size="X-Small" />--%>
                        </asp:GridView>
                        <%-- </ContentTemplate>
                        </asp:UpdatePanel>--%>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="3" style="height: 26px">
                        <asp:Button ID="Button1" runat="server" Text="Button" Visible="False" OnClick="Button1_Click" />
                        <asp:Button ID="btnCancelbottom" runat="server" Text="Cancel" OnClick="btnCanceltop_Click" />&nbsp;
                        <asp:Button ID="btnSubmit" runat="server" OnClick="btnSubmit_Click" Text="Submit"
                            OnClientClick="return validateCAUpdateValue('grp1');" /></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 22px;">
                        <input id="Hidden1" type="hidden" runat="server" />
                        <input id="HdIsOpsTeamMember" type="hidden" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td></td>
                    <td>
                        <asp:Button ID="btnRefresh" runat="server" Text="Refresh" OnClick="btnRefresh_Click"
                            Visible="false" /></td>
                    <td></td>
                </tr>
            </table>


        </div>


        <div class="center">
            <%--<asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Always">

                <ContentTemplate>--%>

            <div id="div1" runat="server" style="flex-item-align: center"></div>



            <input id="Button2" type="button" style="display: none" runat="server" />

            <ajaxToolkit:ModalPopupExtender runat="server"
                ID="performancepopup"
                TargetControlID="Button2"
                PopupControlID="performancepanel"
                BackgroundCssClass="modalBackground"
                DropShadow="true" />



            <br />



            <asp:Panel ID="performancepanel" runat="server" CssClass="modalPopupSecurity" Visible="false">


                <asp:UpdatePanel ID="UpdatePanel3" runat="Server" UpdateMode="Always">
                    <ContentTemplate>
                        <div>
                            <table>
                                <tr>
                                    <td colspan="2" style="text-align: center">
                                        <asp:Label ID="lblTitle" runat="server" Text="PerformanceAddUpdateTool"
                                            Font-Bold="true"></asp:Label>
                                    </td>

                                    <td></td>
                                </tr>
                                <tr>
                                    <td colspan="2">&nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red"
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
                                                    <td align="center" style="font-weight: bold;">Performance As Of Date
                                                    </td>
                                                    <td align="center" style="font-weight: bold;">Performance
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
                                                        <asp:TextBox ID="txtPerformance" runat="server" OnTextChanged="txtPerformance_TextChanged"></asp:TextBox>
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
                                    <td align="right" colspan="2" style="height: 26px">&nbsp;<asp:Button ID="btnsubmitpopup" runat="server" OnClick="btnsubmitpopup_Click" Text="Submit"
                                        OnClientClick="return validateCAUpdateValue('grp1');" />&nbsp;
                        <%--<asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClientClick="return ReturnToParent(null);return false;" OnClick="btnCancel_Click" />--%>&nbsp;
                                                <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="btnCancel_Click" />
                                    </td>
                                </tr>
                                <tr>
                                    <td style="height: 21px"></td>
                                    <td style="height: 21px"></td>
                                </tr>
                                <tr>
                                    <td></td>
                                    <td></td>
                                </tr>
                            </table>
                        </div>

                    </ContentTemplate>


                    <Triggers>
                        <asp:PostBackTrigger ControlID="btnCancel"/>
                        <asp:PostBackTrigger ControlID="btnsubmitpopup" />
                       <%-- <asp:AsyncPostBackTrigger ControlID="btnCancel" EventName="click" />--%>
                    </Triggers>

                </asp:UpdatePanel>
            </asp:Panel>


            <%--  </ContentTemplate>

            </asp:UpdatePanel>--%>
        </div>

    </form>
</body>
</html>
