<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BDPipeline.aspx.cs" Inherits="BDPipeline" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Business Development Pipeline</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <style type="">
 .gvReportss {border-bottom:.02em solid #F2F2F2;}
.ddcblkss {border-bottom:.01em solid #000000;}
.gvReportssNo {border-bottom:.01em solid #ffffff;}
 .gvReportssBlack {border-bottom:.01em solid #000000;}
.ddcblk {border-bottom:.02em solid #F2F2F2;}

.ddcblksswhite {border-bottom:.01em solid #ffffff;}
.BackgroundColor {}

.familyname { font-family:Frutiger 55 Roman;font-size:14pt;font-weight:bold;height:25px; }
.assetdistribution { font-family:Frutiger 55 Roman;font-size:12pt; }
.assDate { font-family:Frutiger 55 Roman;font-size:10pt;font-style:italic; }


     .style2
     {
         width: 4px;
     }


 </style>

    <script type="" language="Javascript">
/**
 * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
    var i;
    for (i = 0; i < s.length; i++){  
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
    var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){  
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
    // February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
    for (var i = 1; i <= n; i++) {
        this[i] = 31
        if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
        if (i==2) {this[i] = 29}
   }
   return this
}

function isDate(dtStr){
    var daysInMonth = DaysArray(12)
    var pos1=dtStr.indexOf(dtCh)
    var pos2=dtStr.indexOf(dtCh,pos1+1)
    var strMonth=dtStr.substring(0,pos1)
    var strDay=dtStr.substring(pos1+1,pos2)
    var strYear=dtStr.substring(pos2+1)
    strYr=strYear
    if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
    if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
    for (var i = 1; i <= 3; i++) {
        if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
    }
    month=parseInt(strMonth)
    day=parseInt(strDay)
    year=parseInt(strYr)
    if (pos1==-1 || pos2==-1){
         
        return false
    }
    if (strMonth.length<1 || month<1 || month>12){
         
        return false
    }
    if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
     
        return false
    }
    if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
     
        return false
    }
    if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
         
        return false
    }
return true
}

function ValidateForm(oSrc, args){
 
     args.IsValid=isDate(args.Value);
   
 }

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <table width="100%">
            <tr>
                <td>
                    <asp:MultiView ID="mvShowReport" runat="server">
                        <asp:View ID="ViewShowFilter" runat="server">
                            <table>
                                <tr>
                                    <td style="width: 100%">
                                        <table width="100%">
                                            <tr>
                                                <td colspan="2" style="width: 20%">
                                                    <img alt="" src="images/Gresham_Logo__.jpg" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2" class="Titlebig" style="width: 20%;">
                                                    Gresham Partners, LLC
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" valign="top" colspan="2">
                                                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    <asp:Label ID="lblContactOwner" runat="server" Text="Contact Owner"></asp:Label></td>
                                                <td style="width: 80%;">
                                                    <asp:DropDownList ID="ddlContactOwner" runat="server" AutoPostBack="True">
                                                    </asp:DropDownList></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    Contact Type</td>
                                                <td style="width: 80%;">
                                                    <asp:ListBox ID="lstContactType" runat="server" Height="195px" SelectionMode="Multiple"
                                                        Width="178px">
                                                        <asp:ListItem Value="0">All</asp:ListItem>
                                                    </asp:ListBox></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    Pipeline</td>
                                                <td style="width: 80%;">
                                                    <asp:DropDownList ID="ddlPipeline" runat="server">
                                                        <asp:ListItem Value="0">Select</asp:ListItem>
                                                        <asp:ListItem>Business Development</asp:ListItem>
                                                        <asp:ListItem>Prospect</asp:ListItem>
                                                        <asp:ListItem>Pre-Prospect</asp:ListItem>
                                                    </asp:DropDownList>
                                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlPipeline"
                                                        Display="None" ErrorMessage="Please select pipeline" InitialValue="0"></asp:RequiredFieldValidator></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    State</td>
                                                <td style="width: 80%;">
                                                    <asp:DropDownList ID="ddlState" runat="server">
                                                    </asp:DropDownList></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    City</td>
                                                <td style="width: 80%;">
                                                    <asp:DropDownList ID="ddlCity" runat="server">
                                                    </asp:DropDownList></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    Type</td>
                                                <td style="width: 80%;">
                                                    <asp:ListBox ID="lstType" runat="server" Height="175px" SelectionMode="Multiple"
                                                        Width="185px">
                                                        <asp:ListItem Value="0">ALL</asp:ListItem>
                                                        <asp:ListItem>Task</asp:ListItem>
                                                        <asp:ListItem>Fax</asp:ListItem>
                                                        <asp:ListItem>Phone Call</asp:ListItem>
                                                        <asp:ListItem>E-mail</asp:ListItem>
                                                        <asp:ListItem>Letter</asp:ListItem>
                                                        <asp:ListItem>Appointment</asp:ListItem>
                                                        <asp:ListItem>Service Activity</asp:ListItem>
                                                        <asp:ListItem>Campaign Response</asp:ListItem>
                                                        <asp:ListItem>Meeting</asp:ListItem>
                                                        <asp:ListItem>Office Meeting</asp:ListItem>
                                                        <asp:ListItem>Schedule A Call</asp:ListItem>
                                                        <asp:ListItem>Meal</asp:ListItem>
                                                        <asp:ListItem>Event</asp:ListItem>
                                                        <asp:ListItem>Correspondence</asp:ListItem>
                                                    </asp:ListBox></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    Firm</td>
                                                <td style="width: 80%;">
                                                    <asp:DropDownList ID="ddlFirms" runat="server">
                                                    </asp:DropDownList></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%" align="left">
                                                    Start Date</td>
                                                <td style="width: 80%">
                                                    <asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>
                                                    <a onclick="showCalendarControl(txtStartDate)">
                                                        <img id="img1" alt="" border="0" src="images/calander.png" style="cursor: hand;" /></a>
                                                    <asp:CustomValidator ID="CustomValidator2" runat="server" ClientValidationFunction="ValidateForm"
                                                        ControlToValidate="txtStartDate" Display="None" ErrorMessage="Start date is not valid"></asp:CustomValidator></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%" align="left">
                                                    End Date</td>
                                                <td style="width: 80%">
                                                    <asp:TextBox ID="txtEndDate" runat="server"></asp:TextBox>
                                                    <a onclick="showCalendarControl(txtEndDate)">
                                                        <img id="img2" alt="" border="0" src="images/calander.png" style="cursor: hand;" /></a>
                                                    <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="ValidateForm"
                                                        ControlToValidate="txtEndDate" Display="None" ErrorMessage="End date is not valid"></asp:CustomValidator>
                                                    <asp:CompareValidator ID="CompareValidator1" runat="server" ErrorMessage="End date can not be less than start date"
                                                        ControlToCompare="txtStartDate" ControlToValidate="txtEndDate" Display="None"
                                                        Operator="GreaterThanEqual" Type="Date"></asp:CompareValidator></td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%" align="left">
                                                    &nbsp;</td>
                                                <td style="width: 80%">
                                                    <asp:RadioButton ID="RadioButton1" runat="server" GroupName="a" Text="HTML" Checked="True" />
                                                    <asp:RadioButton ID="RadioButton2" runat="server" GroupName="a" Text="Excel" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%;" align="left">
                                                    &nbsp;</td>
                                                <td style="width: 80%;">
                                                    <asp:Button ID="Button1" runat="server" Text="Generate Report" OnClick="Button1_Click" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%" align="left">
                                                </td>
                                                <td style="width: 80%">
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td style="width: 20%" align="left">
                                                </td>
                                                <td valign="top" style="width: 80%">
                                                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                                        ShowSummary="False" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </asp:View>
                        <asp:View ID="View1" runat="server">
                            <table width="100%">
                                <tr>
                                    <td>
                                        <table width="100%">
                                            <tr>
                                                <td align="left" style="padding: 5px; height: 35px;">
                                                    <img alt="" height="21" width="144" src="images/Gresham_Logo.jpg" />
                                                </td>
                                                <td align="right" class='noprint' valign="top" style="height: 35px">
                                                    <asp:Button CausesValidation="false" OnClick="btnBack_Click" runat="server" ID="btnBack"
                                                        Text="Back" />
                                                    <asp:Button CausesValidation="false" OnClick="BtnExport_Click" runat="server" ID="GenerateXLS"
                                                        Text="Generate XLS" UseSubmitBehavior="False" />
                                                </td>
                                            </tr>
                                        </table>
                                        <asp:Label ID="lblmessage" runat="server" ForeColor="Red" Visible="False"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td align="center">
                                        <strong>
                                            <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="X-Large" ForeColor="Black">Business Development Pipeline</asp:Label></strong></td>
                                </tr>
                                <tr>
                                    <td align="center">
                                        <asp:Label ID="lblContactName" runat="server" Font-Bold="True" ForeColor="Black"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td align="center">
                                        <asp:Label ID="lblDate" runat="server" Font-Bold="False" Font-Italic="True" ForeColor="Black"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:GridView ID="gvReport" Width="1200px" runat="server" ShowHeader="true" OnRowDataBound="gvReport_RowDataBound">
                                            <Columns>
                                                <asp:BoundField DataField="Contact Type Rank" ItemStyle-Width="100px" HeaderText="Contact Type Rank" />
                                                <asp:BoundField DataField="Contact Name" ItemStyle-Width="100px" HeaderText="Contact Name" />
                                                <%--<asp:BoundField DataField="Contact Type" ItemStyle-Width="100px" HeaderText="Contact Type" />--%>
                                                <asp:BoundField DataField="Firm" HeaderText="Firm" />
                                                <asp:BoundField DataField="Location" HeaderText="Location" />
                                                <asp:BoundField DataField="#" HeaderText="#" />
                                                <asp:BoundField DataField="Date" HeaderText="Date" />
                                                <asp:BoundField DataField="Touchpoint Type - BD Type" HeaderText="Touchpoint Type - BD Type" />
                                                <asp:BoundField DataField="Subject" HeaderText="Subject" />
                                                <asp:BoundField DataField="Owner  Details" HeaderText="Contact Owner/Details" />
                                                <asp:BoundField DataField="_OrderNmb" HeaderText="_OrderNmb" Visible="false" />
                                                <asp:BoundField DataField="_typeflg" HeaderText="_typeflg" Visible="false" />
                                                <asp:BoundField DataField="_BDRankProspectStatus" HeaderText="_BDRankProspectStatus"
                                                    Visible="false" />
                                                <asp:BoundField DataField="_contactname" HeaderText="_contactname" Visible="false" />
                                            </Columns>
                                        </asp:GridView>
                                    </td>
                                </tr>
                            </table>
                        </asp:View>
                    </asp:MultiView>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
