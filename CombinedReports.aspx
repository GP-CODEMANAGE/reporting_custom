<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CombinedReports.aspx.cs" Inherits="CombinedReports" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Meeting Book Schedules</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <style>
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


 </style>

    <script language="Javascript">
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
        <asp:MultiView ID="mvShowReport" runat="server">
            <asp:View ID="ViewShowFilter" runat="server">
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
                                <tr>
                                    <td style="height: 18px" valign="top" colspan="3">
                                        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td style="width: 15%">
                                        <asp:Label ID="lblHouseHold" runat="server" Text="HouseHold:"></asp:Label></td>
                                    <td style="width: 85%">
                                        <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                                        </asp:DropDownList>
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlHousehold"
                                            Display="None" ErrorMessage="Please Select HouseHold"></asp:RequiredFieldValidator>
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 25%">
                                        <asp:Label ID="Label11" runat="server" Text="HouseHold Report Title:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="drpHouseHoldReportTitle" runat="server">
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                               
                                <tr style="display: none;">
                                    <td style="width: 15%;">
                                        <asp:Label ID="Label4" runat="server" Text="Cash :"></asp:Label></td>
                                    <td style="height: 26px">
                                        <asp:DropDownList ID="ddlCash" runat="server">
                                            <asp:ListItem Value="1">Yes</asp:ListItem>
                                            <asp:ListItem Value="0">No</asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px; height: 26px;">
                                    </td>
                                </tr>
                                <tr style="display: none;">
                                    <td style="width: 15%">
                                        <asp:Label ID="Label5" runat="server" Text="Report Group:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="ddlReportFlg" runat="server">
                                            <asp:ListItem Value="1">Yes</asp:ListItem>
                                            <asp:ListItem Value="0">No</asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 15%">
                                        <asp:Label ID="Label10" runat="server" Text="Allocation Group:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="ddlAllocationGroup" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAllocationGroup_SelectedIndexChanged">
                                            <asp:ListItem Value="0">Select</asp:ListItem>
                                        </asp:DropDownList>
                                        </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                 <tr>
                                    <td style="width: 15%">
                                        <asp:Label ID="Label3" runat="server" Text="As Of &nbspDate:"></asp:Label></td>
                                    <td>
                                        <asp:TextBox ID="txtAsofdate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtAsofdate)">
                                            <img id="imgorgDateRec" alt="" border="0" src="images/calander.png" /></a>
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsofdate"
                                            Display="None" ErrorMessage="Please enter As Of Date"></asp:RequiredFieldValidator><asp:CustomValidator
                                                ID="CustomValidator1" runat="server" ControlToValidate="txtAsofdate" ErrorMessage="As of date is not valid"
                                                ClientValidationFunction="ValidateForm" Display="None"> </asp:CustomValidator>
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 15%">
                                        <asp:Label ID="Label12" runat="server" Text="Allocation Group Title:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="drpAllocationGroupTitle" runat="server">
                                            <asp:ListItem Value="0">Select</asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr style="display: none;">
                                    <td style="width: 15%">
                                        <asp:Label ID="Label8" runat="server" Text="Report 1 and 2:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="ddlReport1and2" runat="server">
                                            <asp:ListItem Value="1">Yes</asp:ListItem>
                                            <asp:ListItem Value="0">No</asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr style="display: none;">
                                    <td style="width: 15%">
                                        <asp:Label ID="Label9" runat="server" Text="All Asset :"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="ddlAllAsset" runat="server">
                                            <asp:ListItem Value="1">Yes</asp:ListItem>
                                            <asp:ListItem Value="0">No</asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 15%">
                                        Report:</td>
                                    <td>
                                        <asp:ListBox ID="lstReport" SelectionMode="Multiple" runat="server" Height="125px" Width="237px">
                                        <asp:ListItem Value="0">(All)</asp:ListItem>
                                        <asp:ListItem Value="1">Portfolio Construction Chart</asp:ListItem>
                                        <asp:ListItem Value="2">Commitment Schedule</asp:ListItem>
                                        <asp:ListItem Value="3">Asset Allocation Summary</asp:ListItem>
                                        <asp:ListItem Value="4">Allocation Group Pie Chart</asp:ListItem>
                                        <asp:ListItem Value="5">Overall Pie Chart</asp:ListItem>
                                        </asp:ListBox>
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 15%">
                                    </td>
                                    <td>
                                        <asp:RadioButton ID="RadioButton1" runat="server" Text="HTML" GroupName="a" Visible="False" />&nbsp;<asp:RadioButton
                                            ID="RadioButton2" runat="server" GroupName="a" Text="Excel" Visible="False" />
                                        <asp:RadioButton ID="rdbtnPDF" runat="server" GroupName="a" Text="PDF" Checked="True" /></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                    </td>
                                    <td valign="top">
                                        <asp:Button ID="Button1" runat="server" Text="Generate Report" OnClick="Button1_Click" />
                                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                            ShowSummary="False" />
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:View>
            <asp:View ID="View1" runat="server">
                <table>
                    <tr>
                        <td>
                            <table width="100%">
                                <tr>
                                    <td align="left" style="padding: 5px;">
                                        <img height="21px" width="144px" src="images/Gresham_Logo.png" />
                                    </td>
                                    <td align="right" class='noprint' valign="top">
                                        <asp:Button CausesValidation="false" OnClick="btnBack_Click" runat="server" ID="btnBack"
                                            Text="Back" />
                                        <asp:Button CausesValidation="false" OnClick="BtnExport_Click" runat="server" ID="GenerateXLS"
                                            Text="Generate XLS" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:GridView ID="gvReport" runat="server" OnRowCreated="grd_clientview_onitemcommand">
                            </asp:GridView>
                        </td>
                    </tr>
                </table>
            </asp:View>
        </asp:MultiView>
    </form>
</body>
</html>
