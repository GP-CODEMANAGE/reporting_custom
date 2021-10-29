<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ReportTrackerNew.aspx.cs"
    EnableSessionState="True" Inherits="ReportTrackerNew" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Report Tracker(New)</title>
    <style type="text/css">
        .CellTopBorder {
            border-top-color: Gray;
            border-top: solid;
            border-top-width: thick;
        }
    </style>

    <script type="text/javascript" language="javascript">
        function CheckExtension() {
            var fup = document.getElementById('FileUpload1');
            var fileName = fup.value;
            var ext = fileName.substring(fileName.lastIndexOf('.') + 1);

            if (document.getElementById("ddlAction").value == "11") {
                if (ext != "pdf") {
                    alert("Invalid file please select pdf report to merge.");
                    return false;
                }
            }

        }

        function Uploader() {
            //onchange="Uploader();" 
            document.getElementById('tblBrowse').style.display = "none";
            if (document.getElementById("ddlAction").value == "11") {
                document.getElementById('tblBrowse').style.display = "inline";
                document.getElementById('chkPrepend').checked = false;
                return false;
            }
        }



        function checkAllBoxes() {

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

        function validateCAUpdateValue() {
            var validated = Page_ClientValidate('grp1');
            var frm = document.forms[0];
            //debugger  
            //document.getElementById('ctl00_ContentPlaceHolder1_btnSubmit').disabled = true;  
            if (validated) {

            }
            else {
                alert('Please enter only numeric values in CA Update Value');
                return false;
            }

        }

        function EnableDisable(checkbx, txtbx) {
            if (document.getElementById(checkbx).checked) {
                document.getElementById(txtbx).disabled = true;
                document.getElementById(txtbx).style.backgroundColor = "#EBEBEB"
            }
            else {
                document.getElementById(txtbx).disabled = false;
                document.getElementById(txtbx).style.backgroundColor = "#FFFFFF";
            }

        }

        function SelectAll(id, txtId) {
            var frm = document.forms[0];
            for (i = 0; i < frm.elements.length; i++) {

                if (frm.elements[i].id.match("chkbNC")) {
                    ////debugger;
                    if (document.getElementById(id).checked) {
                        frm.elements[i - 1].innerText = "";
                        frm.elements[i].checked = true;
                        frm.elements[i - 1].disabled = true;
                        frm.elements[i - 1].style.backgroundColor = "#EBEBEB"
                    }
                    else {
                        frm.elements[i - 1].disabled = false;
                        frm.elements[i].checked = false;
                        frm.elements[i - 1].style.backgroundColor = "#FFFFFF"
                    }

                }
            }
        }





    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 100%">
                <tr>
                    <td style="height: 463px">
                        <table width="100%">
                            <tr>
                                <td colspan="2" style="height: 27px">
                                    <img src="images/Gresham_Logo__.jpg" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" class="Titlebig">Gresham Partners, LLC
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 18px" valign="top" colspan="2">
                                    <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                                    <asp:Label ID="noIndex" runat="server" Text="Label" Visible="False"></asp:Label>
                                    <asp:Label ID="chkerror" runat="server" Text="Label" Visible="False"></asp:Label>
                                    <br />
                                    <asp:Label ID="lblError" runat="server" Font-Bold="False" ForeColor="Red" Text="lblError" Visible="False"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana; height: 20px">Batch Type
                                </td>
                                <td style="width: 720px; height: 20px">
                                    <asp:DropDownList ID="ddlBatchType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
                                        <asp:ListItem Value="1">MTGBK</asp:ListItem>
                                        <asp:ListItem Value="4">Merge</asp:ListItem>
                                        <asp:ListItem Value="5">(Q, M)</asp:ListItem>
                                    </asp:DropDownList></td>
                            </tr>

                            <tr>
                                <td style="width: 25%; height: 25px;">
                                    <asp:Label ID="Label1" runat="server" Text="Mail Type" Font-Names="Verdana"></asp:Label></td>
                                <td style="height: 25px">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlMailtype" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlMailtype_SelectedIndexChanged">
                                        <asp:ListItem Value="0">Select</asp:ListItem>
                                        <asp:ListItem Value="1">Capital Call</asp:ListItem>
                                        <asp:ListItem Value="2">Distribution</asp:ListItem>
                                        <asp:ListItem Value="3">Billing</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td style="width: 4px; height: 25px"></td>
                            </tr>
                            <tr>
                                <td style="width: 20%; font-family: Verdana">Batch Owner</td>
                                <td style="width: 720px">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlBatchOwner" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlBatchOwner_SelectedIndexChanged">
                                        <asp:ListItem Value="1">OPS</asp:ListItem>
                                        <asp:ListItem Value="2">Not OPS</asp:ListItem>
                                        <asp:ListItem Value="0">All</asp:ListItem>
                                    </asp:DropDownList>&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana; height: 25px;">Associate</td>
                                <td style="height: 25px; width: 720px;">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlAssociate" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlAssociate_SelectedIndexChanged">
                                    </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana">Household</td>
                                <td style="height: 40px; width: 720px;">
                                    <asp:ListBox ID="lstHouseHold" runat="server" Height="220px" Width="220px" AutoPostBack="True"
                                        OnSelectedIndexChanged="lstHouseHold_SelectedIndexChanged" SelectionMode="Multiple"
                                        Font-Names="Verdana"></asp:ListBox></td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana; height: 21px;">Send Via
                                </td>
                                <td style="height: 21px; width: 720px;">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlSendVia" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlSendVia_SelectedIndexChanged">
                                    </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana; height: 27px;">Report Tracker Status</td>
                                <td style="height: 27px; width: 720px;">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlReportTracker" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlReportTracker_SelectedIndexChanged">
                                    </asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana;">Internal Billing Contact
                                </td>
                                <td style="width: 720px">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlInternalBillingContact" runat="server"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlInternalBillingContact_SelectedIndexChanged">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 25%; font-family: Verdana; height: 20px;">Billing Copy
                                </td>
                                <td style="height: 20px; width: 720px;">
                                    <asp:DropDownList Font-Names="Verdana" ID="ddlBillingCopyHandedOff" runat="server"
                                        AutoPostBack="True" OnSelectedIndexChanged="ddlBillingCopyHandedOff_SelectedIndexChanged">
                                        <asp:ListItem Value="All">All</asp:ListItem>
                                        <asp:ListItem Value="1">Checked</asp:ListItem>
                                        <asp:ListItem Value="0">Unchecked</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <hr />
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="1" valign="top">
                                    <asp:Button ID="btnRefresh" runat="server" Text="Refresh" OnClick="btnRefresh_Click" /></td>
                                <td align="right" colspan="1" valign="top" style="width: 720px">
                                    <table width="100%">
                                        <tr>
                                            <td style="width: 60%" valign="top"></td>
                                            <td align="left" style="width: 40%" valign="top"></td>
                                        </tr>
                                        <tr>
                                            <td style="width: 60%" valign="top">
                                                <table id="tblBrowse" runat="server" width="100%">
                                                    <tr>
                                                        <td>
                                                            <asp:FileUpload ID="FileUpload1" runat="server" /></td>
                                                        <td align="center">
                                                            <asp:CheckBox ID="chkPrepend" runat="server" Text="Prepend" /></td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td align="left" style="width: 40%" valign="top">
                                                <asp:DropDownList Font-Names="Verdana" ID="ddlAction" runat="server" onchange="Uploader();">
                                                    <asp:ListItem Value="1">Approve</asp:ListItem>
                                                    <asp:ListItem Value="4">OPS Change Requested</asp:ListItem>
                                                    <asp:ListItem Value="5">Create Final Report</asp:ListItem>
                                                    <asp:ListItem Value="6">Mark Sent</asp:ListItem>
                                                    <asp:ListItem Value="7">Un-approve</asp:ListItem>
                                                    <asp:ListItem Value="8">Send Billing Copy</asp:ListItem>
                                                    <asp:ListItem Value="9">Remove Hold</asp:ListItem>
                                                    <asp:ListItem Value="10">Update Hold</asp:ListItem>
                                                    <asp:ListItem Value="11">Merge PDF</asp:ListItem>
                                                    <asp:ListItem Value="13">Insert Cover Letter</asp:ListItem>
                                                    <asp:ListItem Value="12">Reject</asp:ListItem>
                                                </asp:DropDownList>
                                                <asp:Button ID="btnSumbitTop" Font-Names="Verdana" runat="server" OnClientClick="return CheckExtension();"
                                                    OnClick="btnSubmit_Click" Text="Submit" /></td>
                                        </tr>
                                        <tr>
                                            <td style="width: 60%" valign="top"></td>
                                            <td align="left" style="width: 40%" valign="top"></td>
                                        </tr>
                                    </table>
                                    &nbsp; &nbsp; &nbsp;&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" style="height: 19px" valign="top">
                                    <asp:Label Font-Names="Verdana" ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large"
                                        Text="Report Tracker"></asp:Label></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td style="width: 100%"></td>
                </tr>
                <tr>
                    <td style="width: 100%;">
                        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="2"
                            BorderWidth="1px" Width="100%" OnRowDataBound="GridView1_RowDataBound" BackColor="White"
                            BorderColor="#CCCCCC" BorderStyle="None" Font-Names="Verdana" Font-Size="X-Small">
                            <Columns>
                                <asp:TemplateField HeaderText="Select">
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelectNC" runat="server" />
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        <input id="chkBoxAll" type="checkbox" onclick="checkAllBoxes()" />
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Batch Type" HeaderText="Batch Type" />
                                <asp:BoundField DataField="Batch Name" HeaderText="Batch Name" />
                                <asp:BoundField DataField="Associate" HeaderText="Associate" />
                                <asp:BoundField DataField="HouseHold" HeaderText="HouseHold" />
                                <asp:BoundField DataField="Recipient" HeaderText="Recipient" />
                                <asp:BoundField DataField="Mailing Status" HeaderText="Mailing Status" />
                                <asp:BoundField DataField="Send VIA" HeaderText="Send VIA" />
                                <asp:BoundField DataField="CA Update" HeaderText="CA Update" />
                                <asp:BoundField DataField="Batch Owner" HeaderText="Batch Owner" />
                                <asp:BoundField DataField="Batch Status" HeaderText="Batch Status" />
                                <asp:BoundField DataField="Reporting Notes" HeaderText="Reporting Notes" />
                                <asp:TemplateField HeaderText="Hold Report">
                                    <ItemTemplate>
                                        <asp:DropDownList runat="server" ID="ddlHoldReport" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Billing Contact" HeaderText="Billing Contact" />
                                <asp:TemplateField HeaderText="Billing Handed Off">
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" ID="chkBillingHandedOff" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Billing Note" HeaderText="Billing Note" />
                                <asp:BoundField DataField="Billing Handed Off" HeaderText="Billing Handed Off" Visible="False" />
                                <asp:BoundField DataField="AdvisorFlag" HeaderText="AadvisorFlag" Visible="False" />
                                <asp:BoundField DataField="ssi_batchid" HeaderText="BatchId" Visible="False" />
                                <asp:BoundField DataField="ssi_mailrecordsid" HeaderText="MailRecordsId" Visible="False" />
                                <asp:BoundField DataField="Advisor Approval" HeaderText="Advisor Approval" Visible="False" />
                                <asp:BoundField DataField="ssi_batchfilename" HeaderText="Batch File Name" Visible="False" />
                                <asp:BoundField DataField="ssi_batchdisplayfilename" HeaderText="Batch Display File Name"
                                    Visible="False" />
                                <asp:TemplateField HeaderText="PDF File">
                                    <ItemTemplate>
                                        <asp:ImageButton runat="server" ID="imgApprovedFile" ImageUrl="~/images/pdf_icon.png"
                                            Height="25px" Width="25px" OnClick="imgApprovedFile_Click" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="ssi_batchdate" HeaderText="batch Date" Visible="False" />
                                <asp:BoundField DataField="ssi_secondaryownerid" HeaderText="ssi_secondaryownerid"
                                    Visible="False" />
                                <asp:BoundField DataField="ssi_reporttrackerstatus" HeaderText="ssi_reporttrackerstatus"
                                    Visible="False" />
                                <asp:BoundField DataField="Handed Off" HeaderText="Handed Off" Visible="False" />
                                <asp:BoundField DataField="FolderNameTxt" HeaderText="FolderName" Visible="False" />
                                <asp:BoundField DataField="As Of Date" HeaderText="As Of Date" Visible="False" />
                                <asp:BoundField DataField="OwnerId" HeaderText="OwnerId" Visible="False" />
                                <asp:BoundField DataField="contactid" HeaderText="contactid" Visible="False" />
                                <asp:BoundField DataField="ssi_holdreport" HeaderText="ssi_holdreport" Visible="False" />
                                <asp:BoundField DataField="hhownerid" HeaderText="hhownerid" Visible="False" />
                                <asp:BoundField DataField="Ssi_InternalBillingContactId" HeaderText="Ssi_InternalBillingContactId"
                                    Visible="False" />
                                <asp:BoundField DataField="Ssi_BillingHandedOff" HeaderText="Ssi_BillingHandedOff"
                                    Visible="False" />
                                <asp:BoundField DataField="ssi_reviewreqdbyid" HeaderText="ssi_reviewreqdbyid" Visible="False" />
                                <asp:BoundField DataField="ssi_reviewreqdbyidname" HeaderText="ssi_reviewreqdbyidname"
                                    Visible="False" />
                                <asp:BoundField DataField="ssi_mailrecordsid" HeaderText="ssi_mailrecordsid" Visible="False" />
                                <asp:BoundField DataField="BatchTypeID" HeaderText="BatchTypeID" Visible="False" />
                                <asp:BoundField DataField="ssi_mailrecords_del" HeaderText="ssi_mailrecords_del"
                                    Visible="False" />
                                <asp:BoundField DataField="BatchType" HeaderText="BatchType" Visible="False" />

                                <asp:BoundField DataField="ssi_billinginvoiceid" HeaderText="ssi_BillingInvoiceId" Visible="False" />
                                <%--28--%>
                            </Columns>
                            <RowStyle ForeColor="#000066" />
                            <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                            <FooterStyle BackColor="White" ForeColor="#000066" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td style="width: 100%">
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                            ShowSummary="False" />
                        <input id="Hidden1" type="hidden" runat="Server" /></td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
