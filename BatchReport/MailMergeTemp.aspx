<%@ Page Theme="Gresham" Debug="true" Language="C#" AutoEventWireup="true" CodeFile="MailMergeTemp.aspx.cs"
    Inherits="_MailMergeTemp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Mail Merge Form Temp</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="../common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="../common/Calendar.js" type="text/javascript"></script>

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

    <script language="Javascript" type="text/javascript">
    
    function ShowHideBrowseButtonRow()
    {
        //document.getElementById("trMonths").style.display = "none";trWireAsof
        document.getElementById("lblError").innerText = "";
        var MailType = document.getElementById("ddlMailType").value;
        // Quarterly/Annual : 0f4c85f4-d0be-e011-a19b-0019b9e7ee05
        // Client Mails : 3bd7d776-e1d3-e011-a19b-0019b9e7ee05
        // General Mails : 99b74584-e2d3-e011-a19b-0019b9e7ee05
        // Smart Mails : c10ba3b7-e1d3-e011-a19b-0019b9e7ee05
        // Prospect Mails : c71108da-e1d3-e011-a19b-0019b9e7ee05 
        
        //trLetter

        
        if(MailType == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || MailType == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || MailType == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || MailType == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || MailType=="" || MailType=="0")
        {
            document.getElementById("trBrowsefiles").style.display = "inline";
            document.getElementById("trReportTemplate").style.display = "inline";
            document.getElementById("trUnify").style.display = "inline";
            document.getElementById("trWireAsof").style.display = "inline";
            document.getElementById("trLetter").style.display = "inline";
        }
        else
        {
            document.getElementById("trBrowsefiles").style.display = "none";
            document.getElementById("trReportTemplate").style.display = "none";
            document.getElementById("trUnify").style.display = "none";
            document.getElementById("trWireAsof").style.display = "none";
            //document.getElementById("trLetter").style.display = "none";
        }
        
        if(MailType == "3fb190d9-b2cd-e011-a19b-0019b9e7ee05")
        {
            document.getElementById("trMonths").style.display = "inline";
        }
        else
        {
            document.getElementById("trMonths").style.display = "none";
        }
        
        return false;
    }
    
    
    function CheckExtension()
    {
       var fup = document.getElementById('FileUpload1');
       var fileName = fup.value;
       var ext = fileName.substring(fileName.lastIndexOf('.') + 1);
       var MailType = document.getElementById("ddlMailType").value;
       var FundName = document.getElementById("lstFund").value;
       
       
       var MailId = document.getElementById("ddlMailId").value;
       var Template = document.getElementById("ddlTemplates").value;
       
       if(MailType == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || MailType == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || MailType == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || MailType == "6d7545da-8164-e111-bd8f-0019b9e7ee05")
       {
            if(MailId == "" || MailId == "0")
            {
                if(Template == "" || Template == "0")
                {
                    alert("Please select Template.");
                    return false;
                }
            }
       }
        
       
      
       
       
//       if(Template == "" || Template == "0")
//       {
//            alert("Please select Template.");
//            return false;
//       }

        // commented 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
//        //debugger;
//       if( MailType == "a1a079a4-d7be-e011-a19b-0019b9e7ee05"       || MailType == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || MailType == "78612b2b-5add-e011-ad4d-0019b9e7ee05"       || MailType== "6d7545da-8164-e111-bd8f-0019b9e7ee05")
//        {
        
//            if(FundName == "" || FundName == "0")
//            {
//                alert("Please select fund.");
//                return false;
//            }
        
        
////            if(ext == "")
////            {
////                return false;
////            }
////        
////            if(ext != "xls")
////            {
////                alert("Please select '.xls' files only.");
////                return false;
////            }
            
//        }
//        else
//        {
            
//        }
        
       
    }
    
    
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
    function ClearLabel()
    {
        document.getElementById("lblError").innerHTML = "";
    }
    
    function selectMonths(posDate) {
        // alert(posDate);
        if (posDate != "") {
            var dt = new Date(posDate);
            var month = dt.getMonth() + 1;
            if (month == 1 || month == 2 || month == 3) {
                document.getElementById('ddlMonths').value = 2;
            }
            else if (month == 4 || month == 5 || month == 6) {
                document.getElementById('ddlMonths').value = 3;
            }
            else if (month == 7 || month == 8 || month == 9) {
                document.getElementById('ddlMonths').value = 4;
            }
            else if (month == 10 || month == 11 || month == 12) {
                document.getElementById('ddlMonths').value = 1;
            }
        }

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
                                        <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
<asp:Label ID="lblErrortxt" runat="server" ForeColor="Red"></asp:Label></td>
                                </tr>
                                 <tr>
                                    <td style="height: 18px" valign="top" colspan="3">
                                        <asp:Label ID="lblSuccess" runat="server" ForeColor="Red"></asp:Label>
                                    </td>
                                   
                                </tr>
                                <tr>
                                    <td colspan="2;">
                                        <asp:LinkButton ID="lbtnExceptionReport" runat="server" OnClick="lbtnExceptionReport_Click" Visible="False">Exception Report</asp:LinkButton>
                                    </td>
                                    
                                </tr>
                                <tr>
                                    <td style="width: 20%">
                                    </td>
                                    <td style="width: 80%">
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trMailID" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="lblMailingId" runat="server" Text="Mail Id:"></asp:Label></td>
                                    <td style="width: 80%">
                                        <asp:DropDownList ID="ddlMailId" runat="server" AutoPostBack="True" 
                                        OnSelectedIndexChanged="ddlMailId_SelectedIndexChanged" onchange="ClearLabel();">
                                        </asp:DropDownList>
                                        <asp:Label ID="lblmailid" runat="server" Text="" Visible="false" ForeColor="Red"></asp:Label>
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 20%; height: 24px;">
                                        <asp:Label ID="lblHouseHold" runat="server" Text="Mailing Type:"></asp:Label></td>
                                    <td style="width: 80%; height: 24px;">
                                        <asp:DropDownList ID="ddlMailType" runat="server" onchange="ShowHideBrowseButtonRow();"
                                            AutoPostBack="True" OnSelectedIndexChanged="ddlMailType_SelectedIndexChanged">
                                        </asp:DropDownList></td>
                                    <td style="width: 4px; height: 24px;">
                                    </td>
                                </tr>
                                <tr id="trReportTemplate" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="lblTemplate" runat="server" Text="Report Template:"></asp:Label></td>
                                    <td style="width: 80%">
                                        <asp:DropDownList ID="ddlTemplates" runat="server" onchange="ClearLabel();">
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trFund" runat="server">
                                    <td style="width: 25%" >
                                        <asp:Label ID="Label11" runat="server" Text="Fund:"></asp:Label></td>
                                    <td style="height: 40px">
                                        <asp:ListBox ID="lstFund" runat="server" onchange="ClearLabel();" Height="150px" Width="300px" ></asp:ListBox></td>
                                    <td style="width: 4px; height: 40px;">
                                    </td>
                                </tr>
                                <tr id="trLegalentity" runat="server">
                                    <td style="width: 25%">
                                    <br />
                                        <asp:Label ID="Label4" runat="server" Text="Legal Entity:"></asp:Label></td>
                                    <td style="height: 40px">
                                     <br />
                                        <asp:ListBox ID="lstLegalEntity" runat="server" onchange="ClearLabel();" Height="150px" SelectionMode="Multiple"></asp:ListBox></td>
                                    <td style="width: 4px; height: 40px;">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 20%">
                                        <asp:Label ID="Label3" runat="server" Text="Position As Of Date:"></asp:Label></td>
                                    <td>
                                        <asp:TextBox ID="txtAsofdate" runat="server"  onChange="selectMonths(this.value);" ></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtAsofdate)">
                                            <img id="imgorgDateRec" runat="server" onclick="ClearLabel();" alt="" border="0" src="images/calander.png" /></a>&nbsp;
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                  <tr id="trAutoDebitDate" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="Label5" runat="server" Text="Auto Debit date:"></asp:Label></td>
                                    <td>
                                        <asp:TextBox ID="txtautoDebitDate" runat="server" ></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtautoDebitDate)">
                                            <img id="img3" runat="server" onclick="ClearLabel();" alt="" border="0" src="images/calander.png" /></a>&nbsp;
                                    </td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trWireAsof" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="Label1" runat="server" Text="Capital Call/Distribution Date:"></asp:Label></td>
                                    <td>
                                        <asp:TextBox ID="txtWireAsofDate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtWireAsofDate)">
                                            <img id="img1" runat="server" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>&nbsp;</td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trLetter" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="Label2" runat="server" Text="Letter Date:"></asp:Label></td>
                                    <td>
                                        <asp:TextBox ID="txtLetterDate" runat="server" ></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtLetterDate)">
                                            <img id="img2" runat="server" alt="" border="0" onclick="ClearLabel();" src="images/calander.png" /></a>&nbsp</td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trBrowsefiles" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="lblUploadFile" runat="server" Text="Upload File:"></asp:Label></td>
                                    <td>
                                        <asp:FileUpload ID="FileUpload1" runat="server" /></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 20%">
                                        <asp:Label ID="lblEmailRecipients" runat="server" Text="Include Email Recipients:"></asp:Label></td>
                                    <td>
                                        <asp:CheckBox ID="chkEmailRecipients" runat="server" onclick="ClearLabel();" /></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trMonths" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="lblMonth" runat="server" Text="Months:"></asp:Label></td>
                                    <td>
                                        <asp:DropDownList ID="ddlMonths" runat="server" onchange="ClearLabel();">
                                            <asp:ListItem Value="1">February – April</asp:ListItem>
                                            <asp:ListItem Value="2">May – July</asp:ListItem>
                                            <asp:ListItem Value="3">August – October</asp:ListItem>
                                            <asp:ListItem Value="4">November – January </asp:ListItem>
                                        </asp:DropDownList></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr id="trUnify" runat="server">
                                    <td style="width: 20%">
                                        <asp:Label ID="lblUnify" runat="server" Text="Unify:"></asp:Label></td>
                                    <td>
                                        <asp:CheckBox ID="chkUnify" runat="server" onclick="ClearLabel();"/></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 20%">
                                    </td>
                                    <td>
                                        <asp:RadioButton ID="RadioButton1" runat="server" Text="HTML" GroupName="a" Visible="False" />&nbsp;<asp:RadioButton
                                            ID="RadioButton2" runat="server" GroupName="a" Text="Excel" Visible="False" />
                                        <asp:RadioButton ID="rdbtnPDF" runat="server" GroupName="a" Text="PDF" Checked="True"
                                            Visible="False" /></td>
                                    <td style="width: 4px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="height: 59px">
                                    </td>
                                    <td valign="top" style="height: 59px">
                                        <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click" OnClientClick="return CheckExtension();return false;" />
                                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                            ShowSummary="False" />
                                    </td>
                                    <td style="width: 4px; height: 59px;">
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:View>
            <asp:View ID="View1" runat="server">
            </asp:View>
        </asp:MultiView>
    </form>
</body>
</html>
