<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MailQueue.aspx.cs" Inherits="MailQueue"
    EnableSessionState="True" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Mail Queue</title>
    <link id="style1" href="../common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="../common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="../common/Calendar.js" type="text/javascript"></script>

    <script language="javascript" type="text/javascript">
        function checkAllBoxes() {

            //get total number of rows in the gridview and do whatever
            //you want with it..just grabbing it just cause
            var totalChkBoxes = parseInt('<%= GridView1.Rows.Count %>');
            var gvControl = document.getElementById('<%= GridView1.ClientID %>');

            //this is the checkbox in the item template...this has to be the same name as the ID of it
            var gvChkBoxControl = "chkSelectNC";

            //this is the checkbox in the header template
            var mainChkBox = document.getElementById("chkBoxAll");

            if (mainChkBox.checked == true) {
                document.getElementById("hdCheckAll").value = "1";
            }
            else if (mainChkBox.checked == false) {
                document.getElementById("hdCheckAll").value = "0";
            }


            //get an array of input types in the gridview
            var inputTypes = gvControl.getElementsByTagName("input");

            for (var i = 0; i < inputTypes.length; i++) {
                //if the input type is a checkbox and the id of it is what we set above
                //then check or uncheck according to the main checkbox in the header template            
                if (inputTypes[i].type == 'checkbox' && inputTypes[i].id.indexOf(gvChkBoxControl, 0) >= 0)
                    inputTypes[i].checked = mainChkBox.checked;
            }
        }


        var TargetBaseControl = null;

        window.onload = function () {
            try {
                //get target base control.
                TargetBaseControl =
                  document.getElementById('<%= this.GridView1.ClientID %>');
        }
        catch (err) {
            TargetBaseControl = null;
        }
    }

    function ValidateCheckBox() {
        if (TargetBaseControl == null) return false;

        //get target child control.
        var TargetChildControl = "chkSelectNC";

        //get all the control of the type INPUT in the base control.
        var Inputs = TargetBaseControl.getElementsByTagName("input");

        for (var n = 0; n < Inputs.length; ++n)
            if (Inputs[n].type == 'checkbox' &&
               Inputs[n].id.indexOf(TargetChildControl, 0) >= 0 &&
               Inputs[n].checked)
                return true;

        alert('Select at least one entry from list!');
        return false;
    }


    function SenderPopUp(MailRecordsId) {


        var WinSettings = "dialogHeight: 500px; dialogWidth: 500px;  edge: Raised; center: Yes; status: no;";
        //var url="Provider.aspx?mode="+ mode + "&pkid=" + ProviderId;
        // $.fn.colorbox({width:"100%", height:"100%", iframe:true, href:url}); 
        var myObject = window.showModalDialog("SenderPopUp.aspx?pkid=" + MailRecordsId, myObject, WinSettings);
        myObject.focus();
        if (myObject != null)
        { __doPostBack('btnRefresh', myObject); }
        else { return false; }
    }

    </script>

</head>
<body>
    <form id="form1" runat="server">
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
                            <td colspan="3" class="Titlebig">Gresham Partners, LLC
                            </td>
                        </tr>
                        <tr>
                            <td class="Titlebig" colspan="3">
                                <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Large"
                                    Text="Report Mail Queue" Width="260px"></asp:Label></td>
                        </tr>
                        <tr>
                            <td style="height: 18px" valign="top" colspan="3"></td>
                        </tr>
                        <tr>
                            <td style="width: 20%">
                                <asp:Label ID="lblType" runat="server" Text="Type:"></asp:Label></td>
                            <td style="width: 80%">
                                <asp:DropDownList ID="ddlType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlType_SelectedIndexChanged">
                                    <asp:ListItem Value="4">Merge Batch Related </asp:ListItem>
                                    <asp:ListItem Value="5">Quarterly/Monthly Batch Related/MTGBK</asp:ListItem>
                                    <asp:ListItem Value="6">Non- Batch Related</asp:ListItem>
                                </asp:DropDownList></td>
                            <td style="width: 4px"></td>
                        </tr>
                        <tr>
                            <td style="width: 20%">
                                <asp:Label ID="lblAsOfDate" runat="server" Text="As Of Date:"></asp:Label></td>
                            <td style="width: 80%">
                                <asp:DropDownList ID="ddlAsofDate" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAsofDate_SelectedIndexChanged">
                                </asp:DropDownList><a onclick="showCalendarControl( txtAsofdate)"></a></td>
                            <td style="width: 4px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%">
                                <asp:Label ID="lblMailId" runat="server" Text="Mail ID:"></asp:Label></td>
                            <td style="height: 40px">
                                <asp:ListBox ID="lstMailId" runat="server" Height="178px" Width="173px" SelectionMode="Multiple"
                                    AutoPostBack="True" OnSelectedIndexChanged="lstMailId_SelectedIndexChanged"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px;"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%">
                                <asp:Label ID="lblMailType" runat="server" Text="Mail Type:"></asp:Label></td>
                            <td style="height: 40px">
                                <asp:ListBox ID="lstMailType" runat="server" Height="200px" AutoPostBack="True" OnSelectedIndexChanged="lstMailType_SelectedIndexChanged"
                                    SelectionMode="Multiple"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 34px;">
                                <asp:Label ID="lblHouseHold" runat="server" Text="Household:"></asp:Label></td>
                            <td style="height: 34px">
                                <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 34px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 33px;">
                                <asp:Label ID="lblAssociate" runat="server" Text="Associate:"></asp:Label></td>
                            <td style="height: 33px">
                                <asp:DropDownList ID="ddlAssociate" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAssociate_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 33px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 29px;">
                                <asp:Label ID="lblAdvisor" runat="server" Text="Advisor:"></asp:Label></td>
                            <td style="height: 29px">
                                <asp:DropDownList ID="ddlAdvisor" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAdvisor_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 29px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%">
                                <asp:Label ID="lblMailPreferance" runat="server" Text="Mail Preference:"></asp:Label></td>
                            <td>
                                <asp:ListBox ID="lstMailPreference" runat="server" Height="183px" Width="196px" SelectionMode="Multiple"
                                    AutoPostBack="True" OnSelectedIndexChanged="lstMailPreference_SelectedIndexChanged"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%">
                                <asp:Label ID="lblMailStatus" runat="server" Text="Mail Status:"></asp:Label></td>
                            <td style="height: 40px">
                                <asp:ListBox ID="lstMailStatus" runat="server" Height="202px" Width="196px" SelectionMode="Multiple"
                                    AutoPostBack="True" OnSelectedIndexChanged="lstMailStatus_SelectedIndexChanged"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 25px;">
                                <asp:Label ID="lblRecipient" runat="server" Text="Salutation Preference:"></asp:Label></td>
                            <td style="height: 25px">
                                <asp:DropDownList ID="ddlSalutationPref" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlSalutationPref_SelectedIndexChanged">
                                    <asp:ListItem Value="1">Both</asp:ListItem>
                                    <asp:ListItem Value="2">Individual</asp:ListItem>
                                    <asp:ListItem Value="3">Joint</asp:ListItem>
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 25px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 28px;">
                                <asp:Label ID="lblCreatedBy" runat="server" Text="Created By:"></asp:Label></td>
                            <td style="height: 28px">
                                <asp:DropDownList ID="ddlCreatedBy" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlCreatedBy_SelectedIndexChanged">
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 28px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 27px;">
                                <asp:Label ID="lblCreatedOn" runat="server" Text="Created On:"></asp:Label></td>
                            <td style="height: 27px">
                                <asp:TextBox ID="txtCreatedOn" runat="server" AutoPostBack="True" OnTextChanged="txtCreatedOn_TextChanged"
                                    CausesValidation="True"></asp:TextBox>&nbsp; <a onclick="showCalendarControl( txtCreatedOn)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" /></a>
                                <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="txtCreatedOn"
                                    ErrorMessage="Invalid Created On Date" ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$">*</asp:RegularExpressionValidator></td>
                            <td style="width: 4px; height: 27px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 27px"></td>
                            <td style="height: 27px"></td>
                            <td style="width: 4px; height: 27px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 27px">
                                <asp:Label ID="lblTest" runat="server" Text="Test Flag:"></asp:Label></td>
                            <td style="height: 27px">
                                <asp:CheckBox ID="chkTest" runat="server" /></td>
                            <td style="width: 4px; height: 27px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%; height: 27px"></td>
                            <td style="height: 27px"></td>
                            <td style="width: 4px; height: 27px"></td>
                        </tr>
                        <tr>
                            <td style="width: 25%"></td>
                            <td>
                                <asp:Button ID="btnSerach" runat="server" OnClick="btnSerach_Click" Text="Search"
                                    Visible="False" /></td>
                            <td style="width: 4px;"></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td valign="top">
                                <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                    ShowSummary="False" DisplayMode="List" />
                            </td>
                            <td style="width: 4px;"></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                               <div id="divControlContainer" runat="server" style="flex-item-align:center"></div>
                                    <asp:Label ID="lblError3" runat="server" ForeColor="Black"></asp:Label>
                                <%-- </div>--%>
                            </td>
                            <td style="width: 4px"></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:Label ID="lblError2" runat="server" ForeColor="Red"></asp:Label><br />
                                <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>

                                 <div id="divControlContainer2" runat="server" style="flex-item-align:center"></div>
                                
                            </td>
                            <td style="width: 4px"></td>
                        </tr>
                        <tr>
                            <td align="right" style="text-align: center" colspan="2" valign="top">
                                <hr />
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnRefresh" runat="server" Text="Refresh" OnClick="btnRefresh_Click" /></td>
                            <td align="right" colspan="1" valign="top">
                                <asp:DropDownList ID="ddlAction" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAction_SelectedIndexChanged">
                                    <asp:ListItem Value="1">Save Reports to Sharepoint and Client Portal</asp:ListItem>
                                    <asp:ListItem Value="11">Save Reports to SharePoint only</asp:ListItem>
                                    <asp:ListItem Value="2">Create Mailing CSV for Merge</asp:ListItem>
                                    <asp:ListItem Value="3">Create Single PDF for Batch Reports</asp:ListItem>
                                    <asp:ListItem Value="4">Create FED-EX CSV</asp:ListItem>
                                    <asp:ListItem Value="5">Mark All Records Canceled</asp:ListItem>
                                    <asp:ListItem Value="6">Mark All Records Sent</asp:ListItem>
                                    <asp:ListItem Value="7">Send Email to Associate and Mark Sent</asp:ListItem>
                                    <asp:ListItem Value="8">Mark Printed</asp:ListItem>
                                    <asp:ListItem Value="9">Create Individual PDFs –Grouped by Household and Recipient</asp:ListItem>
                                    <asp:ListItem Value="10">Insert Mailing Sheet</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClick="btnSubmit_Click" />
                                &nbsp;
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td align="right" colspan="2" valign="top">
                                <table width="35%">
                                    <tr>
                                        <td style="width: 15%; height: 21px">
                                            <asp:Label ID="lblOptions" runat="server" Font-Bold="True" Text="Options:"></asp:Label></td>
                                        <td align="left" style="height: 21px">
                                            <asp:CheckBox ID="chkGroupRecandSpouse" runat="server" Text="Group Recipient and Spouse" /></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 15%"></td>
                                        <td align="left">
                                            <asp:CheckBox ID="chkMailingSheets" runat="server" Text="Exclude Mailing Sheets" /></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 15%"></td>
                                        <td align="left">
                                            <asp:CheckBox ID="chkReportSeperator" runat="server" Text="Exclude Report Seperator Sheets" /></td>
                                    </tr>
                                </table>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td colspan="2" valign="top">
                                <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                    DataKeyNames="ssi_mailrecordsId" CellPadding="2" AllowSorting="True" OnSorting="GridView1_Sorting"
                                    OnRowDataBound="GridView1_RowDataBound" BackColor="White" BorderColor="#CCCCCC"
                                    BorderStyle="None" BorderWidth="1px" Font-Names="Verdana" Font-Size="X-Small">
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="chkSelectNC" />
                                            </ItemTemplate>
                                            <HeaderTemplate>
                                                <input id="chkBoxAll" type="checkbox" onclick="checkAllBoxes()" />
                                            </HeaderTemplate>
                                        </asp:TemplateField>
                                        <%--0--%>
                                        <asp:BoundField HeaderText="Mail ID" DataField="Mail ID" />
                                        <%--1--%>
                                        <asp:BoundField HeaderText="Mail Type" DataField="Mail Type" SortExpression="Mail Type" />
                                        <asp:BoundField HeaderText="As Of Date" DataField="As Of Date" SortExpression="As Of Date" />
                                        <asp:BoundField HeaderText="Household/Related Batch" DataField="Related Batch" />
                                        <asp:BoundField HeaderText="Mail Status" DataField="Mail Status" />
                                        <%--5--%>
                                        <asp:BoundField HeaderText="Recipient" DataField="Receipent" SortExpression="Receipent" />
                                        <asp:BoundField HeaderText="Mailing Address/Email" DataField="Mailing Address/ Email" />
                                        <asp:BoundField HeaderText="Mail Preference" DataField="Mail Preference" SortExpression="Mail Preference" />
                                        <asp:BoundField HeaderText="Salutation Preference" DataField="Salutation Preference" />
                                        <asp:BoundField HeaderText="Created By" DataField="CreatedBy" />
                                        <%--10--%>
                                        <asp:BoundField HeaderText="Created On" DataField="Createdon" />
                                        <asp:BoundField HeaderText="BatchId" DataField="ssi_batchId" Visible="False" />
                                        <asp:BoundField HeaderText="MailingRecordsId" DataField="ssi_mailrecordsId" Visible="False" />
                                        <asp:BoundField DataField="FolderNameTxt" HeaderText="FolderName" Visible="False" />
                                        <asp:BoundField DataField="HouseholdNameTxt" HeaderText="HouseholdNameTxt" Visible="False" />
                                        <%--15--%>
                                        <asp:BoundField DataField="PdfFileName" HeaderText="PdfFileName" Visible="False" />
                                        <asp:BoundField DataField="ssi_mailrecordsId" HeaderText="ssi_mailrecordsId" Visible="False" />
                                        <asp:BoundField DataField="ssi_batchfilename" HeaderText="PdfFilePath" Visible="False" />
                                        <asp:BoundField DataField="Ssi_ClientPortalFolder" HeaderText="Ssi_ClientPortalFolder" Visible="False" />
                                        <asp:BoundField DataField="Ssi_SharePointReportFolder" HeaderText="Ssi_SharePointReportFolder" Visible="False" />
                                        <%--20--%>
                                        <asp:BoundField DataField="ssi_batchdisplayfilename" HeaderText="ssi_batchdisplayfilename" Visible="False" />
                                        <asp:BoundField DataField="ReviewReqdBy" HeaderText="ReviewReqdBy" />
                                        <asp:BoundField DataField="Ssi_ReviewReqdById" HeaderText="Ssi_ReviewReqdById" Visible="False" />
                                        <asp:BoundField DataField="Ssi_reporttrackerstatus" HeaderText="Ssi_reporttrackerstatus" Visible="False" />
                                        <asp:BoundField DataField="Mail StatusID" HeaderText="Mail StatusID" Visible="False" />
                                        <%--25--%>
                                        <asp:BoundField DataField="ssi_secondaryownerid" HeaderText="ssi_secondaryownerid" Visible="False" />
                                        <asp:BoundField DataField="ssi_holdreport" HeaderText="ssi_holdreport" Visible="False" />
                                        <asp:BoundField DataField="ssi_BillingHandedOff" HeaderText="ssi_BillingHandedOff" Visible="False" />
                                        <asp:BoundField DataField="ssi_internalbillingcontactid" HeaderText="ssi_internalbillingcontactid" Visible="False" />
                                        <asp:BoundField DataField="Contact Address/ Email" HeaderText="Contact Address/ Email" Visible="False" />
                                        <%--30--%>
                                        <asp:BoundField DataField="BatchTypeID" HeaderText="BatchTypeID" Visible="False" />
                                        <asp:BoundField DataField="SubfolderName" HeaderText="SubfolderName" Visible="False" />
                                        <asp:BoundField DataField="ssi_spvfilename" HeaderText="SubfolderName" Visible="False" />
                                        <asp:BoundField DataField="ssi_salutation_mail" HeaderText="ssi_salutation_mail" Visible="False" />
                                        <asp:BoundField DataField="FirstNameSort" HeaderText="FirstNameSort" Visible="False" />
                                        <%--35--%>
                                        <asp:BoundField DataField="LastNameSort" HeaderText="LastNameSort" Visible="False" />
                                        <asp:BoundField DataField="BatchType" HeaderText="BatchType" Visible="False" />
                                        <asp:BoundField DataField="Batch_Name" HeaderText="Batch_Name" Visible="False" />
                                        <asp:BoundField DataField="ssi_clientportalname" HeaderText="ssi_clientportalname" Visible="False" />
                                         <asp:BoundField DataField="ssi_CSSiteUUID" HeaderText="ssi_CSSiteUUID" Visible="False" />
                                          <%--40--%> 
                                        <asp:BoundField DataField="ssi_CSLegalEntityUUID" HeaderText="ssi_CSLegalEntityUUID" Visible="False" />
                                         <asp:BoundField DataField="ssi_SPLEFolder" HeaderText="ssi_SPLEFolder" Visible="False" />
                                         <asp:BoundField DataField="ssi_SPSiteType" HeaderText="ssi_SPSiteType" Visible="False" />
                                         <%--43--%>
                                         <asp:BoundField DataField="ssi_billinginvoiceid" HeaderText="ssi_billinginvoiceid" Visible="False" />  <%--44--%>
                                         <asp:BoundField DataField="ssi_billingid" HeaderText="ssi_billingid" Visible="False" /> <%--45--%>

                                    </Columns>
                                    <FooterStyle BackColor="White" ForeColor="#000066" />
                                    <RowStyle ForeColor="#000066" />
                                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                </asp:GridView>
                            </td>
                            <td>
                                <input id="Hidden1" type="hidden" runat="Server" />
                                <asp:HiddenField ID="hdCheckAll" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
