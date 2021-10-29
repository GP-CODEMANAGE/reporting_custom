<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OPSUpdateForm.aspx.cs" Inherits="OPSUpdateForm" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>OPS Update form</title>
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

        .center {
            height: 450px;
            width: 919px;
        }
    </style>

    <script type="text/javascript" language="javascript">
        // For Auto refreshing the grid values
        function Refressh() {
            //debugger;
            if (event.keyCode == 13) {
                __doPostBack("txtCashinTransit", "TextChanged");
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
                alert('Please enter only numeric values in Ops Update Value');
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


        function SelectAll1(id, txtId) {
            //get reference of GridView control
            var grid = document.getElementById("<%= gvList.ClientID %>");
            //variable to contain the cell of the grid
            var cell;

            if (grid.rows.length > 0) {
                //loop starts from 1. rows[0] points to the header.
                for (i = 1; i < grid.rows.length; i++) {
                    //get the reference of first column
                    cell = grid.rows[i].cells[12];
                    cell2 = grid.rows[i].cells[11];
                    //loop according to the number of childNodes in the cell
                    for (j = 0; j < cell.childNodes.length; j++) {
                        //if childNode type is CheckBox                 
                        if (cell.childNodes[j].type == "checkbox") {
                            //assign the status of the Select All checkbox to the cell checkbox within the grid
                            cell.childNodes[j].checked = document.getElementById(id).checked;

                        }
                        if (cell2.childNodes[j].type == "input") {
                            cell2.childNodes[j].style.backgroundColor = "#FFFFFF"
                        }
                    }
                }
            }
        }


    </script>

    <script language="javascript" type="text/javascript">

        function OpenChild(PositionId) {
            var WinSettings = "dialogHeight: 175px; dialogWidth: 775px;  edge: Raised; center: Yes; status: no;";
            var myObject = window.showModalDialog("PopUpOpsUpdateSolution.aspx?posid=" + PositionId, myObject, WinSettings);

            if (myObject != null) {
                __doPostBack('btnRefresh', myObject);
                return false;
            }
            else {
                return false;
            }
        }

        function ClearLabel() {
            document.getElementById("lblMessage").innerHTML = "";
        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <div>
            <table width="100%">
                <tr>
                    <td style="width: 204px;">&nbsp;
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan="3">
                        <strong>Secondary Owner &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:DropDownList
                            ID="ddlSecOwner" runat="server" OnSelectedIndexChanged="ddlSecOwner_SelectedIndexChanged"
                            AutoPostBack="True">
                        </asp:DropDownList></strong></td>
                </tr>
                <tr>
                    <td colspan="3">
                        <b>Household </b>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
                        &nbsp; &nbsp;&nbsp;
                        <asp:DropDownList ID="ddlHousehold" AutoPostBack="true" runat="server" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td colspan="3">
                        <b>Update Month</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        :
                        <asp:Label ID="lblUpdateMonth" Font-Bold="true" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 22px">
                        <input id="Hidden1" type="hidden" runat="server" />
                        <input id="hidHouseholdId" type="hidden" />
                        <asp:CheckBox ID="chkbxNCAll" runat="server" Text="N/C" OnCheckedChanged="chkbNCAll_CheckedChanged"
                            AutoPostBack="True" Visible="False" /></td>
                </tr>
                <tr>
                    <td align="left" colspan="3" style="height: 22px">
                        <asp:Label ID="lblMessage" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" colspan="3" style="height: 22px;">&nbsp;<asp:DropDownList ID="ddlReportOptions" runat="server" onchange="ClearLabel();">
                        <asp:ListItem Selected="True" Value="0">Update</asp:ListItem>
                        <asp:ListItem Value="1">Run PDF</asp:ListItem>
                        <asp:ListItem Value="2">Run and Update PDF</asp:ListItem>
                        <asp:ListItem Value="3">Approve Household Reports</asp:ListItem>
                    </asp:DropDownList>
                        <asp:Button ID="btnSumbitTop" runat="server" OnClick="btnSubmit_Click" Text="Submit"
                            OnClientClick="return validateCAUpdateValue('grp1');" Style="height: 26px" />&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 50px;">
                        <asp:Label ID="lblClientSpecific" runat="server" Font-Bold="True" ForeColor="DarkOrange"
                            Text="CLIENT SPECIFIC UPDATES"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 232px">
                        <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="False" TabIndex="1"
                            ToolTip="OPS Update" AllowSorting="True" OnRowDataBound="gvList_RowDataBound" OnRowCommand="gvList_RowCommand"
                            Font-Names="Verdana" Font-Size="X-Small" CellPadding="2">
                            <Columns>
                                <asp:BoundField DataField="ssi_positionid" HeaderText="ssi_positionid" Visible="False" />
                                <asp:BoundField DataField="ssi_accountid" HeaderText="AccountId" Visible="False" />
                                <asp:BoundField DataField="Account" HeaderText="Account" />
                                <asp:BoundField DataField="Name1" HeaderText="Name1">
                                    <ItemStyle Width="300px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Name2" HeaderText="Name 2">
                                    <ItemStyle Width="400px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="SecSymbol" HeaderText="Sec Symbol" />
                                <asp:BoundField DataField="Security" HeaderText="Security">
                                    <ItemStyle Width="250px" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Commitment">
                                    <ItemTemplate>
                                        <asp:LinkButton runat="server" ID="lnkedit" class="link" Text="Edit commitment" CommandName="linkButton1" CommandArgument="<%# Container.DataItemIndex %>"></asp:LinkButton>
                                        <%--<a runat="server" id="lnkedit" class="link">Edit Commitment</a>--%>
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        Commitment
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="AssetClass" HeaderText="Asset Class" />
                                <asp:BoundField DataField="datasource" HeaderText="Source" />
                                <asp:BoundField DataField="ModifiedOn" DataFormatString="{0:MM/dd/yyyy}" HeaderText="Last Update on Account" />
                                <asp:BoundField DataField="ModifiedbyName" HeaderText="Last Updated By" />
                                <asp:BoundField DataField="PreviousMonthMktValue" DataFormatString="{0:$#,###0;($#,###0)}"
                                    HeaderText="Previous Month Value">
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="CurrentMonthMktValue" DataFormatString="{0:$#,###0;($#,###0)}"
                                    HeaderText="Current Month Value">
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="OPS Update Value">
                                    <ItemTemplate>
                                        <asp:TextBox runat="server" ID="txtCAUpdateValue" Width="82px" /><asp:RegularExpressionValidator
                                            ID="RegularExpressionValidator1" runat="server" ControlToValidate="txtCAUpdateValue"
                                            Display="Dynamic" ErrorMessage="Please enter numeric values only" ValidationExpression="^-?\d*(\.\d+)?$"
                                            ValidationGroup="grp1">*</asp:RegularExpressionValidator>
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        OPS Update Value
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="N/C">
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" ID="chkbNC" />
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        N/C<br />
                                        <asp:CheckBox ID="chkbxNCSelectAll" runat="server" />
                                    </HeaderTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Commitment" HeaderText="Commitment" Visible="False" />
                                <asp:BoundField DataField="sas_assetclassid" Visible="False"></asp:BoundField>
                                <asp:BoundField DataField="Ssi_subassetclassId" Visible="False"></asp:BoundField>
                                <asp:BoundField DataField="Ssi_BenchmarkSubAssetClassId" Visible="False"></asp:BoundField>
                                <asp:BoundField DataField="SectorFlg" Visible="False"></asp:BoundField>
                                <asp:BoundField DataField="Ssi_LoadLockDT" Visible="False"></asp:BoundField>
                                <asp:BoundField DataField="_UpdateFlg" HeaderStyle-CssClass="displayNone" ControlStyle-CssClass="displayNone" ItemStyle-CssClass="displayNone"></asp:BoundField>
                            </Columns>
                            <HeaderStyle Height="10px" BackColor="#BFDBFF" Font-Size="X-Small" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="3" style="height: 26px">
                        <asp:Button ID="Button1" runat="server" Text="Button" Visible="False" />
                        <asp:Button ID="btnSubmit" runat="server" OnClick="btnSubmit_Click" Text="Submit"
                            OnClientClick="return validateCAUpdateValue('grp1');" /></td>
                </tr>
                <tr>
                    <td style="width: 204px; height: 21px;"></td>
                    <td style="height: 21px"></td>
                    <td style="height: 21px"></td>
                </tr>
                <tr>
                    <td style="width: 204px"></td>
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
                        <table>
                            <tr>
                                <td></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td colspan="2">&nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" valign="top">
                                    <asp:GridView ID="gvPopUp" runat="server" AutoGenerateColumns="False" TabIndex="1"
                                        ToolTip="CA Update" AllowSorting="True" OnRowDataBound="gvPopUp_RowDataBound"
                                        Font-Names="Verdana" Font-Size="X-Small" CellPadding="2">
                                        <Columns>
                                            <asp:BoundField DataField="Security" HeaderText="Security" ItemStyle-Width="200px" />
                                            <asp:BoundField DataField="Last Update" HeaderText="Last Update" ItemStyle-Width="150px">
                                                <ItemStyle />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Last Updated By" HeaderText="Last Updated By" ItemStyle-Width="200px">
                                                <ItemStyle />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="PreviousMonthMktValue" HeaderText="Previous Month Commitment"
                                                ItemStyle-Width="200px" ItemStyle-HorizontalAlign="Right" />
                                            <asp:BoundField DataField="CurrentMonthMktValue" HeaderText="Current Month Commitment"
                                                ItemStyle-Width="200px" ItemStyle-HorizontalAlign="Right"></asp:BoundField>
                                            <asp:TemplateField HeaderText="OPS Update Value">
                                                <ItemTemplate>
                                                    <asp:TextBox runat="server" ID="txtCAUpdateValue" Width="100px" /><asp:RegularExpressionValidator
                                                        ID="RegularExpressionValidator1" runat="server" ControlToValidate="txtCAUpdateValue"
                                                        Display="Dynamic" ErrorMessage="Please enter numeric values only" ValidationExpression="^-?\d*(\.\d+)?$"
                                                        ValidationGroup="grp1">*</asp:RegularExpressionValidator>
                                                </ItemTemplate>
                                                <HeaderTemplate>
                                                    OPS Update Value
                                                </HeaderTemplate>
                                            </asp:TemplateField>
                                        </Columns>
                                        <HeaderStyle Height="10px" BackColor="#BFDBFF" Font-Size="X-Small" />
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" colspan="2" style="height: 26px">&nbsp;<asp:Button ID="btnsubmitpopup" runat="server" OnClick="btnsubmitpopup_Click" Text="Submit"
                                    OnClientClick="return validateCAUpdateValue('grp1');" />&nbsp;
                        <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClientClick="return ReturnToParent(null);return false;" OnClick="btnCancel_Click" />&nbsp;
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

                    </ContentTemplate>


                    <Triggers>
                        <asp:PostBackTrigger ControlID="btnCancel" />
                        <asp:PostBackTrigger ControlID="btnsubmitpopup" />
                    </Triggers>

                </asp:UpdatePanel>
            </asp:Panel>


            <%--  </ContentTemplate>

            </asp:UpdatePanel>--%>
        </div>
    </form>
</body>
</html>
