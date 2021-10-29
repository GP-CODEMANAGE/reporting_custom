<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AttributionComparison.aspx.cs"
    Inherits="AttributionComparison" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Attribution Comparison</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">

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
        function ValidateDate() {
            //debugger;
            //    var Household=document.getElementById("ddlHousehold").value;

            //    var AllocGroup=document.getElementById("ddlAllocationGroup").value;
            //    var Year=document.getElementById("ddlYear").value;
            var AsofDate = document.getElementById("txtAsOfDate").value;

            //    if(Household=="0" || Household=="")
            //    {
            //        alert("Please Select Household.");
            //        return false;
            //    }

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
                //        AsofDate=AsofDate.split('/');
                //        if(AsofDate[2]!=Year)
                //        {
                //            alert("Please select the same end date as the year selected.");
                //            return false;
                //        }
            }
            //document.getElementById("trSubmit").style.display = "none";
        }
        function ClearLabel() {
            document.getElementById("lblMessage").innerHTML = "";
            document.getElementById("trDownLoad").style.display = "none";
        }

        function ClearDropdown(ddl) {
            //debugger;
            //document.getElementById(ddl).value="0";
            ClearLabel();
        }

        function FreeFormDate() {
            //debugger;
            var AsofDate = document.getElementById("txtAsOfDate").value;

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

                document.getElementById("txtAsOfDate").value = mon + "/" + day + "/" + year;
            }
            else {
                //                 alert("Please provide proper M/D/YYYY or M/D/YY or MM/dd/yyyy format date.");
                //                 return false;
            }

        }

        function isGoodDate(dt) {
            var reGoodDate = /^(?:(0[1-9]|1[012])[\/.](0[1-9]|[12][0-9]|3[01])[\/.](19|20)[0-9]{2})$/;
            return reGoodDate.test(dt);
        }
    </script>
    <style type="text/css">
        .style1
        {
            width: 199px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width: 100%">
            <tr>
                <td>
                    <table style="width: 100%">
                        <tr>
                            <td colspan="3">
                                <img src="images/Gresham_Logo__.jpg" />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" class="Titlebig">
                                Gresham Partners, LLC
                            </td>
                        </tr>
                        <tr style="height: 50px;">
                            <td colspan="3" align="left">
                                <asp:Label ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large" Text="Attribution Comparison"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" colspan="3">
                                <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                        <td>
                        Groups
                        </td>
                        <td>
                          <asp:DropDownList  ID="ddlGroups" runat="server" AutoPostBack="True" 
                                Font-Names="Verdana">
                                    <asp:ListItem Value="9">Select</asp:ListItem>
                                    <asp:ListItem Value="1">Family GA</asp:ListItem>
                                    <asp:ListItem Value="0">Composite</asp:ListItem>
                                </asp:DropDownList>&nbsp;
                        </td>
                        </tr>
                        <tr>
                            <td class="style1">
                                <asp:Label ID="Label4" runat="server" Text="As of Date:"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="txtAsOfDate" runat="server" onchange="FreeFormDate();" Text=""></asp:TextBox>&nbsp;&nbsp;<a
                                    onclick="showCalendarControl( txtAsOfDate)">
                                    <img id="imgorgDateRec" alt="" onclick="ClearLabel();" border="0" src="images/calander.png" /></a>
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr id="trSubmit" runat="server">
                            <td class="style1">
                            </td>
                            <td valign="top">
                                <br />
                                <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClientClick="return ValidateDate();"
                                    OnClick="btnSubmit_Click" />
                                <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                    ShowSummary="False" />
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr runat="server" id="trDownLoad">
                            <td class="style1">
                            </td>
                            <td valign="top">
                            </td>
                            <td>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
