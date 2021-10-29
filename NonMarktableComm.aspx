<%@ Page Theme="Gresham" Debug="true" Language="C#" AutoEventWireup="true" CodeFile="NonMarktableComm.aspx.cs"
    Inherits="NonMarktableComm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Non-Marketable Commitments</title>
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


     .style2
     {
         width: 4px;
     }


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
        <table width="100%">
            <tr>
                <td style="height: 19px">
                    <asp:MultiView ID="mvShowReport" runat="server">
                        <asp:View ID="ViewShowFilter" runat="server">
                            <table>
                                <tr>
                                    <td>
                                        <table style="width: 388px">
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
                                                <td>
                                                    Advisor</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlAdvisor" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAdvisor_SelectedIndexChanged">
                                                    </asp:DropDownList></td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    Associate</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlAssociate" runat="server">
                                                    </asp:DropDownList></td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="lblHouseHold" runat="server" Text="Close Date"></asp:Label></td>
                                                <td>
                                                    <asp:DropDownList ID="ddlclosedate" runat="server">
                                                    </asp:DropDownList>
                                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlclosedate"
                                                        Display="Dynamic" ErrorMessage="Please Select a Close Date" InitialValue="0">*</asp:RequiredFieldValidator></td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="lblType" runat="server" Text="Type"></asp:Label></td>
                                                <td>
                                                    <asp:ListBox ID="lstType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="lstType_SelectedIndexChanged" SelectionMode="Multiple"></asp:ListBox></td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="Label11" runat="server" Text="Partnership"></asp:Label></td>
                                                <td>
                                                    <asp:ListBox ID="lstpartnership" runat="server" SelectionMode="Multiple" Rows="10"></asp:ListBox>
                                                </td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:RadioButton ID="RadioButton1" runat="server" Checked="True" GroupName="a" Text="HTML" />
                                                    <asp:RadioButton ID="RadioButton2" runat="server" GroupName="a" Text="Excel" />
                                                </td>
                                                <td class="style2">
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Generate Report" />
                                                    <br />
                                                </td>
                                                <td class="style2">
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged"
                                                        Visible="False">
                                                    </asp:DropDownList>
                                                </td>
                                                <td>
                                                    <asp:TextBox ID="txtpriorperiod" runat="server" Visible="False"></asp:TextBox>&nbsp;<a
                                                        onclick="showCalendarControl(txtpriorperiod)">
                                                        <img id="img1" alt="" border="0" src="images/calander.png" height="0" width="0" /></a>
                                                    <asp:CustomValidator ID="CustomValidator2" runat="server" ControlToValidate="txtpriorperiod"
                                                        ErrorMessage="Prior Period Comparison is not valid" ClientValidationFunction="ValidateForm"
                                                        Display="None" Visible="False"></asp:CustomValidator></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="drpHouseHoldReportTitle" runat="server" Visible="False">
                                                    </asp:DropDownList>
                                                </td>
                                                <td>
                                                    <asp:TextBox ID="txtAsofdate" runat="server" Visible="False"></asp:TextBox>&nbsp;&nbsp;<a
                                                        onclick="showCalendarControl( txtAsofdate)">
                                                        <img id="imgorgDateRec" alt="" border="0" height="0" width="0" src="images/calander.png" /></a>
                                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsofdate"
                                                        Display="None" ErrorMessage="Please enter as of date" Visible="False"></asp:RequiredFieldValidator><asp:CustomValidator
                                                            ID="CustomValidator1" runat="server" ControlToValidate="txtAsofdate" ErrorMessage="As of date is not valid"
                                                            ClientValidationFunction="ValidateForm" Display="None" Visible="False"></asp:CustomValidator>
                                                </td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlLookthrough" runat="server" Visible="False">
                                                        <asp:ListItem>Detail</asp:ListItem>
                                                        <asp:ListItem>Consolidated</asp:ListItem>
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlContact" runat="server" Visible="False">
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlAllocationGroup" runat="server" AutoPostBack="True" OnSelectedIndexChanged="drpAllocationGroupTitle_SelectedIndexChanged"
                                                        Visible="False">
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="drpAllocationGroupTitle" runat="server" Visible="False">
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="drpSummary" runat="server" Visible="False">
                                                        <asp:ListItem>Detail column</asp:ListItem>
                                                        <asp:ListItem>Summary column</asp:ListItem>
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlVersion" runat="server" Visible="False">
                                                        <asp:ListItem>Yes</asp:ListItem>
                                                        <asp:ListItem>No</asp:ListItem>
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlReportGroupflag" runat="server" Visible="False">
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    <asp:DropDownList ID="ddlReportgroupflag2" runat="server" Visible="False">
                                                    </asp:DropDownList></td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                </td>
                                                <td>
                                                    <asp:DropDownList ID="ddlAlignment" runat="server" Visible="False">
                                                        <asp:ListItem>Horizontal</asp:ListItem>
                                                        <asp:ListItem>Vertical</asp:ListItem>
                                                    </asp:DropDownList></td>
                                                <td class="style2">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                </td>
                                                <td>
                                                    &nbsp;</td>
                                                <td style="width: 4px">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                </td>
                                                <td valign="top">
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
                                                    <img height="45" width="114" src="images/Gresham_Logo.png" />
                                                </td>
                                                <td align="right" class='noprint' valign="top">
                                                    <asp:Button CausesValidation="false" OnClick="btnBack_Click" runat="server" ID="btnBack"
                                                        Text="Back" />
                                                    <asp:Button CausesValidation="false" OnClick="BtnExport_Click" runat="server" ID="GenerateXLS"
                                                        Text="Generate XLS" />
                                                </td>
                                            </tr>
                                        </table>
                                        <asp:Label ID="lblmessage" runat="server" Visible="False"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td align="center" style="font-size: 20px">
                                        <strong>Non-Marketable Commitments</strong></td>
                                </tr>
                                <tr>
                                    <td align="center" style="font-size: 20px">
                                        <asp:Label ID="lblPartnership" runat="server" Font-Bold="True" Font-Size="20px"></asp:Label></td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:GridView ID="gvReport" runat="server" OnRowCreated="gvReport_RowCreated" OnRowDataBound="gvReport_RowDataBound">
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
