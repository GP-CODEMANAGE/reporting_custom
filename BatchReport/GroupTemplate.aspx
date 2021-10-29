<%@ Page Theme="Gresham" Debug="true" Language="C#" AutoEventWireup="true" CodeFile="GroupTemplate.aspx.cs"
    Inherits="_GroupTemplate" Culture="en-US" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Group Template</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <script language="Javascript" type="text/javascript">
    
        function validategroup()
        {
           // if(document.getElementById("") == 
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
 

   
    </script>

    <script language="javascript" type="text/javascript">
 
        function ClearLabel()
        {
            document.getElementById("lblError").innerHTML = "";
        }
 
    </script>
    <style type="text/css">
        .auto-style1 {
            height: 18px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table style="width: 100%">
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
                <td valign="top" colspan="3" class="auto-style1">
                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
          <%--  <tr id="trMessage" runat="server">
                <td align="center" colspan="2">
                    <asp:Label ID="Label1" runat="server" Text="Label" ForeColor="Red" Visible="False"></asp:Label>
                </td>
                <td style="width: 4px"></td>
            </tr>--%>
            <tr>
                <td>
                     <asp:LinkButton ID="lbtnExceptionReport" runat="server" Text="Exception Report" OnClick="lbtnExceptionReport_Click" Visible="False" ></asp:LinkButton>
                </td>
                </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate1" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate2" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate3" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%; height: 26px;"></td>
                <td style="height: 26px">
                    <asp:DropDownList ID="ddlGroupTemplate4" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px; height: 26px;"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate5" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate6" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate7" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate8" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate9" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate10" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate11" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate12" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate13" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate14" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:DropDownList ID="ddlGroupTemplate15" runat="server" onchange="ClearLabel();">
                    </asp:DropDownList></td>
                <td style="width: 4px"></td>
            </tr>
            <tr>
                <td style="width: 20%">
                    <asp:Label ID="lblUnify" runat="server" Text="Unify:"></asp:Label>
                </td>
                <td>
                    <asp:CheckBox ID="chkUnify" runat="server" onclick="ClearLabel();" /></td>
            </tr>
            <tr>
                <td style="width: 20%"></td>
                <td>
                    <asp:RadioButton ID="RadioButton1" runat="server" Text="HTML" GroupName="a" Visible="False" />&nbsp;<asp:RadioButton
                        ID="RadioButton2" runat="server" GroupName="a" Text="Excel" Visible="False" />
                    <asp:RadioButton ID="rdbtnPDF" runat="server" GroupName="a" Text="PDF" Checked="True"
                        Visible="False" /></td>
                <td style="width: 4px"></td>
            </tr>
            <tr id="trbtnSubmit" runat="server">
                <td></td>
                <td valign="top" align="left">&nbsp;<asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click" />&nbsp;
                </td>
                <td style="width: 4px"></td>
            </tr>
        </table>
    </form>
</body>
</html>
