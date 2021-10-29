<%@ Page Theme="Gresham" Debug="true" Language="C#" AutoEventWireup="true" CodeFile="ReportReviewForm.aspx.cs"
    MaintainScrollPositionOnPostback="true" Inherits="_ReportReviewForm" Culture="en-US" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Report Review</title>
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

    <script language="Javascript" type="text/javascript">
    function CheckExtension()
    {
        var fup = document.getElementById('FileUpload1');
        var fileName = fup.value;
        var ext = fileName.substring(fileName.lastIndexOf('.') + 1);
        
        if(document.getElementById("ddlAction").value == "8")
        {
            if(ext != "pdf" )
            {
                alert("Invalid file please select pdf report to merge.");
                return false;
            }
        }

    }
    
    function Uploader()
    {
    //onchange="Uploader();" 
     document.getElementById('tblBrowse').style.display = "none"; 
       if(document.getElementById("ddlAction").value == "8")
       {
            document.getElementById('tblBrowse').style.display = "inline"; 
            document.getElementById('chkPrepend').checked = false;
           return false;
       }
    }
      function checkAllBoxes(){

  //get total number of rows in the gridview and do whatever
  //you want with it..just grabbing it just cause
  var totalChkBoxes = parseInt('<%= GridView1.Rows.Count %>');   
  var gvControl = document.getElementById('<%= GridView1.ClientID %>');
           
  //this is the checkbox in the item template...this has to be the same name as the ID of it
  var gvChkBoxControl = "chkSelectNC";  
           
  //this is the checkbox in the header template
  var mainChkBox = document.getElementById("chkBoxAll");
           
  //get an array of input types in the gridview
  var inputTypes = gvControl.getElementsByTagName("input");
           
  for(var i = 0; i < inputTypes.length; i++)
  {  
     //if the input type is a checkbox and the id of it is what we set above
     //then check or uncheck according to the main checkbox in the header template            
     if(inputTypes[i].type == 'checkbox' && inputTypes[i].id.indexOf(gvChkBoxControl,0) >= 0)
          inputTypes[i].checked = mainChkBox.checked;  
  }
} 
    
    
function openEmail()
{
	mail_str ='mailto:abc@yahoo.com?subject=More info here in link:?attachment="C:\\Reports\\Rahul_20111114_191951\\Anathan Family_2011-0930.pdf"';
	mail_str += "&body= You should look at this report by clicking the below link:%0D%0A" + escape("<a href='http://yahoo.com'>click here</a>"); 
	location.href = mail_str;
}
    
    
function ClearMessage()
{
    var ddlView = document.getElementById("ddlView").value;
    var ddlAdvisor = document.getElementById("ddlAdvisor").value;
    var ddlAssociate = document.getElementById("ddlAssociate").value;
    
    if(ddlView != "" || ddlAdvisor !="" || ddlAssociate!= "")
    {
    
    }
}
    
    
    function ShowHideBrowseButtonRow()
    {
        document.getElementById("lblError").innerText = "";
        var MailType = document.getElementById("ddlMailType").value;
        
        // Quarterly/Annual : 0f4c85f4-d0be-e011-a19b-0019b9e7ee05
        // Client Mails : 3bd7d776-e1d3-e011-a19b-0019b9e7ee05
        // General Mails : 99b74584-e2d3-e011-a19b-0019b9e7ee05
        // Smart Mails : c10ba3b7-e1d3-e011-a19b-0019b9e7ee05
        // Prospect Mails : c71108da-e1d3-e011-a19b-0019b9e7ee05
        
        if(MailType == "0f4c85f4-d0be-e011-a19b-0019b9e7ee05" || MailType == "3bd7d776-e1d3-e011-a19b-0019b9e7ee05" || MailType == "99b74584-e2d3-e011-a19b-0019b9e7ee05" || MailType == "c10ba3b7-e1d3-e011-a19b-0019b9e7ee05" || MailType == "c71108da-e1d3-e011-a19b-0019b9e7ee05")
        {
            document.getElementById("trBrowsefiles").style.display = "none";
        }
        else
        {
            //document.getElementById("trBrowsefiles").style.display = "none";
            document.getElementById("trBrowsefiles").style.display = "inline";
        }
        
        return false;
    }
    
    

    
    
    
var TargetBaseControl = null;
        
   window.onload = function()
   {
      try
      {
         //get target base control.
         TargetBaseControl = 
           document.getElementById('<%= this.GridView1.ClientID %>');
      }
      catch(err)
      {
         TargetBaseControl = null;
      }
   }
        
   function ValidateCheckBox()
   {              
      if(TargetBaseControl == null) return false;
      
      //get target child control.
      var TargetChildControl = "chkSelectNC";
            
      //get all the control of the type INPUT in the base control.
      var Inputs = TargetBaseControl.getElementsByTagName("input"); 
            
      for(var n = 0; n < Inputs.length; ++n)
         if(Inputs[n].type == 'checkbox' && 
            Inputs[n].id.indexOf(TargetChildControl,0) >= 0 && 
            Inputs[n].checked)
          return true;        
            
      alert('Select at least one entry from list!');
      return false;
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

</head>
<body>
    <form id="form1" runat="server">
        <table style="width: 100%">
            <a href="http://crm01/ISV/AdventReport/BatchReport/ReportReviewForm.aspx?id={Batch GUID(Batch)}">
            </a>
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
                            <td class="Titlebig" colspan="3">
                                <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Large"
                                    Text="Report Review" Width="260px"></asp:Label></td>
                        </tr>
                        <tr>
                            <td style="height: 18px" valign="top" colspan="3">
                                <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                                <asp:Label ID="noIndex" runat="server" Text="Label" Visible="False"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 18px" valign="top">
                                <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr>
                            <td style="width: 20%">
                                <asp:Label ID="lblView" runat="server" Text="View:" Font-Names="Verdana"></asp:Label></td>
                            <td style="width: 80%">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlView" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlView_SelectedIndexChanged">
                                    <asp:ListItem Value="0">All</asp:ListItem>
                                    <asp:ListItem Value="1">Action Required By Me</asp:ListItem>
                                </asp:DropDownList>&nbsp;
                            </td>
                            <td style="width: 4px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 24px;">
                                <asp:Label ID="lblAdvisor" runat="server" Text="Advisor:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 24px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlAdvisor" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlAdvisor_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 24px;">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 25px;">
                                <asp:Label ID="lblAssociate" runat="server" Text="Associate:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 25px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlAssociate" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlAssociate_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 25px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%">
                                <asp:Label ID="lblHouseHold" runat="server" Text="Household:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 40px">
                                <asp:ListBox Font-Names="Verdana" ID="lstHouseHold" runat="server" Height="220px"
                                    Width="220px" AutoPostBack="True" OnSelectedIndexChanged="lstHouseHold_SelectedIndexChanged"
                                    SelectionMode="Multiple"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 25px;">
                                <asp:Label ID="lblBatchType" runat="server" Text="Batch Type:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 25px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlBatchtype" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlBatchtype_SelectedIndexChanged">
                                    <asp:ListItem Value="0">All</asp:ListItem>
                                    <asp:ListItem Value="2">Quarterly</asp:ListItem>
                                    <asp:ListItem Value="3">Monthly</asp:ListItem>
                                    <asp:ListItem Value="4">Merge</asp:ListItem>
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 25px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 22px;">
                                <asp:Label ID="lblBatchOwner" runat="server" Text="Batch Owner:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 22px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlBatchOwner" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlBatchOwner_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 22px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 26px;">
                                <asp:Label ID="lblBatchStatus" runat="server" Text="Batch Status:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 26px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlBatchstatus" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlBatchstatus_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 26px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 17px;">
                                <asp:Label ID="lblMailStatus" runat="server" Text="Mail Status:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 17px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlMailStatus" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlMailStatus_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 17px">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 25px;">
                                <asp:Label ID="lblRecipient" runat="server" Text="Recipient:" Font-Names="Verdana"></asp:Label></td>
                            <td style="height: 25px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlRecipient" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlRecipient_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 25px">
                            </td>
                        </tr>
                        <tr>
                            <td style="border-bottom: gray 1px solid; text-align: center; height: 12px;" colspan="2">
                                &nbsp;
                            </td>
                            <td style="width: 4px; height: 12px;">
                            </td>
                        </tr>
                        <tr>
                            <td align="left" colspan="1" style="height: 21px" valign="middle">
                                <asp:Button ID="btnRefresh" runat="server" Text="Refresh" OnClick="btnRefresh_Click" /></td>
                            <td align="right" colspan="1" valign="middle" style="height: 21px">
                                <table width="100%">
                                    <tr>
                                        <td style="width: 50%">
                                        </td>
                                        <td align="left" style="width: 30%">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 50%" valign="top">
                                            <table id="tblBrowse" runat="server" width="100%">
                                                <tr>
                                                    <td>
                        <asp:FileUpload ID="FileUpload1" runat="server" /></td>
                        <td align="center">
                            <asp:CheckBox ID="chkPrepend" runat="server" Text="Prepend" /></td>
                                                </tr>
                                            </table>
                                        </td>
                                        <td align="left" style="width: 30%" valign="top">
                                <asp:DropDownList Font-Names="Verdana"  ID="ddlAction" runat="server" onchange="Uploader();">
                                    <asp:ListItem Value="2">Review PDF/Batch</asp:ListItem>
                                    <asp:ListItem Value="1">Approve</asp:ListItem>
                             <%--       <asp:ListItem Value="2">Review PDF/Batch</asp:ListItem>--%>
                                    <asp:ListItem Value="3">Request OPS Change</asp:ListItem>
                                    <asp:ListItem Value="4">Unapprove/Remove PDF</asp:ListItem>
                                    <asp:ListItem Value="5">Remove Hold</asp:ListItem>
                                    <asp:ListItem Value="6">Update Hold</asp:ListItem>
                                    <asp:ListItem Value="7">Billing Complete</asp:ListItem>
                                    <asp:ListItem Value="8">Merge PDF</asp:ListItem>
                                    <asp:ListItem Value="10">Insert Cover Letter</asp:ListItem>
                                     <asp:ListItem Value="9">Reject</asp:ListItem>
                                </asp:DropDownList>
                                            <asp:Button Font-Names="Verdana" ID="btnSubmit" runat="server" Text="Submit" OnClick="btnSubmit_Click" OnClientClick="return CheckExtension();" /></td>
                                    </tr>
                                </table>
                                &nbsp; &nbsp;<%--   <asp:Button ID="Button1" runat="server" Visible="false" Text="Submit1" OnClientClick="openEmail();"
                                    OnClick="btnSubmit_Click"OnClientClick="javascript:return ValidateCheckBox();" />--%>&nbsp;
                            </td>
                            <td style="width: 4px; height: 21px">
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" valign="top">
                                <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                    OnRowDataBound="GridView1_RowDataBound" BackColor="White" BorderColor="#CCCCCC"
                                    Font-Names="Verdana" Font-Size="X-Small" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="2">
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="chkSelectNC" />
                                            </ItemTemplate>
                                            <HeaderTemplate>
                                                <input id="chkBoxAll" type="checkbox" onclick="checkAllBoxes()" />
                                            </HeaderTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField HeaderText="Batch Name" DataField="Batch Name" />
                                        <asp:BoundField HeaderText="Email Or Mailing Address" DataField="Email Or Mailing Address" />
                                        <asp:BoundField HeaderText="Batch Owner" DataField="Batch Owner" />
                                        <asp:BoundField HeaderText="Batch Status" DataField="Batch Status" />
                                        <asp:BoundField HeaderText="Next Batch Owner" DataField="Next Batch Owner" />
                                        <asp:BoundField HeaderText="Mailing Status" DataField="Mailing Status" />
                                        <asp:BoundField HeaderText="Send Via" DataField="Send Via" />
                                        <asp:TemplateField HeaderText="Hold Report">
                                            <ItemTemplate>
                                                <asp:DropDownList runat="server" ID="ddlHoldReport" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Billing Handed Off">
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="chkBillingHandedOff" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField HeaderText="BatchId" DataField="ssi_batchid" Visible="False" />
                                        <asp:BoundField HeaderText="Billing Handed Off" DataField="Billing Handed Off" Visible="False" />
                                        <asp:BoundField HeaderText="AadvisorFlag" DataField="AdvisorFlag" Visible="False" />
                                        <asp:TemplateField HeaderText="Approved File">
                                            <ItemTemplate>
                                                <asp:ImageButton runat="server" ID="imgApprovedFile" ImageUrl="~/images/pdf_icon.png"
                                                    Height="25px" Width="25px" OnClick="imgApprovedFile_Click" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="FolderNameTxt" HeaderText="FolderName" Visible="False" />
                                        <asp:BoundField DataField="HouseholdNameTxt" HeaderText="HouseholdNameTxt" Visible="False" />
                                        <asp:BoundField DataField="PdfFileName" HeaderText="PdfFileName" Visible="False" />
                                        <asp:BoundField DataField="ssi_batchfilename" HeaderText="Batch File Name" Visible="False" />
                                        <asp:BoundField DataField="ssi_batchdisplayfilename" HeaderText="Batch Display File Name"
                                            Visible="False" />
                                        <asp:BoundField DataField="BatchStatusID" HeaderText="BatchStatusID" Visible="False" />
                                        <asp:BoundField DataField="OwnerId" HeaderText="Batch Owner ID" Visible="False" />
                                        <asp:BoundField DataField="ssi_secondaryownerid" HeaderText="ssi_secondaryownerid"
                                            Visible="False" />
                                        <asp:BoundField DataField="ssi_holdreport" HeaderText="ssi_holdreport" Visible="False" />
                                        <%--<asp:BoundField DataField="OwnerId1" HeaderText="HouseHold Owner ID" Visible="False" />--%>
                                        <asp:BoundField DataField="BatchOwnerId" HeaderText="BatchOwnerId" Visible="False" />
                                        <asp:BoundField DataField="MailStatusId" HeaderText="MailStatusId" Visible="False" />
           <asp:BoundField DataField="ssi_mailrecordsid" HeaderText="ssi_mailrecordsid" Visible="False" />
              <asp:BoundField DataField="BatchTypeID" HeaderText="BatchTypeID" Visible="False" />
              <asp:BoundField DataField="ssi_mailrecords_del" HeaderText="ssi_mailrecords_del" Visible="False" />
                                    </Columns>
                                    <FooterStyle BackColor="White" ForeColor="#000066" />
                                    <RowStyle ForeColor="#000066" />
                                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                </asp:GridView>
                            </td>
                            <td style="width: 4px; height: 59px">
                            </td>
                        </tr>
                    </table>
                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                        ShowSummary="False" />
                    &nbsp;
                    <input id="Hidden1" type="hidden" runat="Server" />&nbsp;
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
