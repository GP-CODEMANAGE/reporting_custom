<%@ Page Language="C#" AutoEventWireup="true" CodeFile="TestTemplate.aspx.cs" Inherits="TestTemplate"
    EnableEventValidation="false" MaintainScrollPositionOnPostback="true" %>

<%@ Register TagPrefix="cc" Namespace="Winthusiasm.HtmlEditor" Assembly="Winthusiasm.HtmlEditor" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Create Template</title>
    <link id="style1" href="common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="common/Calendar.js" type="text/javascript"></script>

    <script language="javascript" type="text/javascript">

        function Remove(ddlFundType, ddlFund, trFund) {
            //code done on server side commented on 06-28-2013

            //        document.getElementById(trFund).style.display = "none";
            //        document.getElementById(ddlFundType).value="0";
            //        document.getElementById(ddlFund).value = "0";
            //        document.getElementById("hdtblId").value = "1";
            //return false;
        }

        function getFundvalues() {
            var str1 = "1";
            if (document.getElementById("rdoYes").checked == true) {
                document.getElementById("trFund1").style.display = "inline";
                document.getElementById("trAddAnother").style.display = "inline";

                var tblval = parseInt(document.getElementById("hdtblId").value) + 1;
                if (tblval < 11) {
                    //document.getElementById("hdtblId").value = tblval; hdFunds
                    //document.getElementById("trFund" + tblval).style.display = "inline";
                    //debugger;
                    for (var i = 1; i < 11; i++) {
                        var strFundType = document.getElementById("ddlFundType" + i).value;
                        var strFund = document.getElementById("ddlFund" + i).value;


                        if (strFund != "" && strFund != "0") {
                            str1 = str1 + "|" + strFund;
                        }
                    }

                    //get all selected funds
                    if (str1 != "") {
                        document.getElementById("hdFunds").value = str1.substring(2, str1.length);
                        return false;
                    }

                    //return false;
                }

            }
        }

        function GetHtmlEditor() {
            return $find('<%= txtFundDesc1.ClientID %>');
            return false;
        }


        function AddAnotherFund() {

            //document.getElementById("hdtblId").value="1";
            // debugger;
            if (document.getElementById("rdoYes").checked == true) {
                //document.getElementById("hdtblId").value =1;
                document.getElementById("trFund1").style.display = "inline";
                document.getElementById("trAddAnother").style.display = "inline";

                var tblval = parseInt(document.getElementById("hdtblId").value) + 1;
                if (tblval < 11) {
                    document.getElementById("hdtblId").value = tblval;
                    document.getElementById("trFund" + tblval).visible = true;
                }
                return false;
            }
        }


        function ResetFundSpecificValues() {
            //debugger;
            if (document.getElementById("rdoYes").checked == true) {
                __doPostBack('btnAddFund', 'btnAddFund_Click');
                return false;
            }
            else if (document.getElementById("rdoNo").checked == true) {
                __doPostBack('btnAddFund', 'btnAddFund_Click');
                //document.getElementById("hdtblId").value = "1";
                return false;
            }
        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePartialRendering="true">
        </asp:ScriptManager>
        <div>
            <asp:HiddenField ID="hdtblId" runat="server" Value="1" />
            <asp:HiddenField ID="hdFunds" runat="server" />
            <asp:HiddenField ID="hdEditRows" runat="server" />
            <table width="85%">
                <tr>
                    <td style="width: 25%">
                        <img src="Images/Gresham_Logo__.jpg" /></td>
                    <td style="width: 60%"></td>
                </tr>
                <tr>
                    <td style="width: 25%">Gresham Partners, LLC
                    </td>
                    <td style="width: 60%"></td>
                </tr>
                <tr>
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <asp:Label ID="lblError" runat="server" Text="" ForeColor="Red">
                        </asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="lblTemplateFund" runat="server" Text="Template:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:DropDownList ID="ddlTemplate" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlTemplate_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="lblTemplateName" runat="server" Text="Template Name:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:TextBox ID="txtTemplate" runat="server">
                        </asp:TextBox></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="lblTemplateType" runat="server" Text="Template Type:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:DropDownList ID="ddlTemplateType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlTemplateType_SelectedIndexChanged">
                        </asp:DropDownList></td>
                </tr>
                <tr id="trFileupload" runat="server">
                    <td style="width: 25%">
                        <asp:Label ID="lblFile" runat="server" Text="File:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:FileUpload ID="FileUpload1" runat="server" /></td>
                </tr>
                <tr runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <asp:Label ID="lblFileName" runat="server" Text="Label"></asp:Label></td>
                </tr>
                <tr>
                    <td style="width: 25%;">
                        <asp:Label ID="lblAsOfDate" runat="server" Text="As Of Date:"></asp:Label></td>
                    <td style="width: 60%;">
                        <asp:TextBox ID="txtAsOfDate" runat="server">
                        </asp:TextBox>
                        <a onclick="showCalendarControl(txtAsOfDate)">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="lblLetterDate" runat="server" Text="Letter Date:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:TextBox ID="txtDateOfLetter" runat="server">
                        </asp:TextBox>&nbsp;<a onclick="showCalendarControl(txtDateOfLetter)">
                            <img id="imgorgDateRec" alt="" border="0" src="images/calander.png" /></a></td>
                </tr>
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="lblLetterText" runat="server" Text="Letter Text:"></asp:Label></td>
                    <td style="width: 60%"></td>
                </tr>
                <tr>
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <div>
                            <%--<asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                <ContentTemplate>
                                    <cc:HtmlEditor ID="txtLetterText" runat="server" Height="200px" Width="400px" />
                                </ContentTemplate>
                            </asp:UpdatePanel>--%>
                            <FCKeditorV2:FCKeditor ID="txtLetterText" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                BasePath="~/FCKeditor/" Width="60%" Height="250px">
                            </FCKeditorV2:FCKeditor>
                            &nbsp;&nbsp;&nbsp;
                        </div>
                    </td>
                </tr>
                <tr id="trFundSpecific" runat="server">
                    <td style="width: 25%">
                        <asp:Label ID="lblIncludeFund" runat="server" Text="Include Fund Specific Sections:">
                        </asp:Label></td>
                    <td style="width: 60%">
                        <asp:RadioButton ID="rdoYes" runat="server" Checked="True" GroupName="FundSpecific"
                            onclick="ResetFundSpecificValues();" Text="Yes" AutoPostBack="True" OnCheckedChanged="rdoYes_CheckedChanged" />&nbsp;<asp:RadioButton
                                ID="rdoNo" onclick="ResetFundSpecificValues();" runat="server" GroupName="FundSpecific"
                                Text="No" /></td>
                </tr>
                <tr id="trFund1" runat="server">
                    <td style="width: 25%;"></td>
                    <td style="width: 60%;">
                        <table id="tblFund1" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                                <td align="left"></td>
                            </tr>

                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType1_SelectedIndexChanged">
                                            </asp:DropDownList></td>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>

                                    <asp:UpdatePanel ID="UpdatePanel21" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund1" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td align="left"></td>
                            </tr>


                            <tr>
                                <td colspan="2" valign="top">
                                    <%--<asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                        <ContentTemplate>
                                            <cc:HtmlEditor ID="txtFundDesc1" runat="server" EnableViewState="true" Height="200px"
                                                Width="400px" />
                                        </ContentTemplate>
                                    </asp:UpdatePanel>--%>
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc1" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="60%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">

                                    <asp:LinkButton ID="lnkRemove1" runat="server" OnClientClick="return Remove('ddlFundType1','ddlFund1','trFund1');"
                                        OnClick="lnkRemove1_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                        &nbsp;&nbsp;
                    </td>
                </tr>
                <tr id="trFund2" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="tblFund2" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                                <td align="left"></td>
                            </tr>

                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType2" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType2_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel22" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund2" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td align="left"></td>
                            </tr>

                            <tr>
                                <td colspan="2" valign="top">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc2" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove2" runat="server" OnClientClick="return Remove('ddlFundType2','ddlFund2','trFund2');return false;"
                                        OnClick="lnkRemove2_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                        &nbsp;&nbsp;
                    </td>
                </tr>
                <tr id="trFund3" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="tblFund3" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                                <td align="left"></td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel23" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType3" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType3_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel3" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund3" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>


                                </td>
                                <td align="left"></td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc3" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove3" runat="server" OnClientClick="return Remove('ddlFundType3','ddlFund3','trFund3');return false;"
                                        OnClick="lnkRemove3_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                        &nbsp;&nbsp;
                    </td>
                </tr>
                <tr id="trFund4" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table1" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel24" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType4" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType4_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel4" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund4" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc4" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove4" runat="server" OnClientClick="return Remove('ddlFundType4','ddlFund4','trFund4');return false;"
                                        OnClick="lnkRemove4_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund5" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table2" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel25" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType5" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType5_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel5" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund5" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc5" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove5" runat="server" OnClientClick="return Remove('ddlFundType5','ddlFund5','trFund5');return false;"
                                        OnClick="lnkRemove5_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund6" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table3" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel26" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType6" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType6_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel6" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund6" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc6" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove6" runat="server" OnClientClick="return Remove('ddlFundType6','ddlFund6','trFund6');return false;"
                                        OnClick="lnkRemove6_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund7" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table4" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel27" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType7" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType7_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel7" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund7" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc7" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove7" runat="server" OnClientClick="return Remove('ddlFundType7','ddlFund7','trFund7');return false;"
                                        OnClick="lnkRemove7_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund8" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table5" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel28" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType8" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType8_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel8" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund8" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc8" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove8" runat="server" OnClientClick="return Remove('ddlFundType8','ddlFund8','trFund8');return false;"
                                        OnClick="lnkRemove8_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund9" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table6" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel29" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType9" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFundType9_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel9" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund9" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc9" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="250px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove9" runat="server" OnClientClick="return Remove('ddlFundType9','ddlFund9','trFund9');return false;"
                                        OnClick="lnkRemove9_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund10" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table7" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel30" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType10" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType10_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel10" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund10" runat="server">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc10" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove10" runat="server" OnClientClick="return Remove('ddlFundType10','ddlFund10','trFund10');return false;"
                                        OnClick="lnkRemove10_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund11" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table8" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel31" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType11" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType11_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel11" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund11" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc11" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove11" runat="server" OnClientClick="return Remove('ddlFundType11','ddlFund11','trFund11');return false;"
                                        OnClick="lnkRemove11_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund12" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table9" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel32" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType12" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType12_SelectedIndexChanged">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel12" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund12" runat="server">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc12" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove12" runat="server" OnClientClick="return Remove('ddlFundType12','ddlFund12','trFund12');return false;"
                                        OnClick="lnkRemove12_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund13" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table10" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel33" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType13" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType13_SelectedIndexChanged">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel13" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund13" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc13" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove13" runat="server" OnClientClick="return Remove('ddlFundType13','ddlFund13','trFund13');return false;"
                                        OnClick="lnkRemove13_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund14" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table11" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel34" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType14" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType14_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel14" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund14" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc14" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove14" runat="server" OnClientClick="return Remove('ddlFundType14','ddlFund14','trFund14');return false;"
                                        OnClick="lnkRemove14_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund15" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table12" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel35" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType15" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType15_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel15" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund15" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc15" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove15" runat="server" OnClientClick="return Remove('ddlFundType15','ddlFund15','trFund15');return false;"
                                        OnClick="lnkRemove15_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund16" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table13" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel36" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType16" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType16_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel16" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund16" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc16" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove16" runat="server" OnClientClick="return Remove('ddlFundType16','ddlFund16','trFund16');return false;"
                                        OnClick="lnkRemove16_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund17" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table14" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel37" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType17" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType17_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel17" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund17" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>


                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc17" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove17" runat="server" OnClientClick="return Remove('ddlFundType17','ddlFund17','trFund17');return false;"
                                        OnClick="lnkRemove17_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund18" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table15" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel38" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType18" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType18_SelectedIndexChanged">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel18" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund18" runat="server">
                                            </asp:DropDownList>

                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc18" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove18" runat="server" OnClientClick="return Remove('ddlFundType18','ddlFund18','trFund18');return false;"
                                        OnClick="lnkRemove18_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund19" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table16" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel39" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType19" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType19_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel19" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund19" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc19" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove19" runat="server" OnClientClick="return Remove('ddlFundType19','ddlFund19','trFund19');return false;"
                                        OnClick="lnkRemove19_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trFund20" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <table id="Table17" runat="server" width="100%">
                            <tr>
                                <td>Fund Type</td>
                                <td>Fund</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel40" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFundType20" runat="server" EnableViewState="true" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlFundType20_SelectedIndexChanged">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>

                                </td>
                                <td>
                                    <asp:UpdatePanel ID="UpdatePanel20" runat="server" UpdateMode="Always">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="ddlFund20" runat="server">
                                            </asp:DropDownList>
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <FCKeditorV2:FCKeditor ID="txtFundDesc20" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                        BasePath="~/FCKeditor/" Width="65%" Height="200px">
                                    </FCKeditorV2:FCKeditor>
                                </td>
                                <td colspan="1" align="left">
                                    <asp:LinkButton ID="lnkRemove20" runat="server" OnClientClick="return Remove('ddlFundType20','ddlFund20','trFund20');return false;"
                                        OnClick="lnkRemove20_Click">Remove</asp:LinkButton>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="trAddAnother" runat="server">
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <asp:Button ID="btnAddFund" runat="server" Text="Add Another Fund" OnClick="btnAddFund_Click" /></td>
                </tr>
                <tr id="trDynamic" runat="server">
                    <td style="width: 25%">
                        <asp:Label ID="lblDynamic" runat="server" Text="Dynamic:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:RadioButton ID="rdoDynamic" runat="server" GroupName="Dynamic" Text="Yes" AutoPostBack="True" Checked="True" OnCheckedChanged="rdoDynamic_CheckedChanged" />
                        <asp:RadioButton ID="rdoStatic" runat="server" GroupName="Dynamic" Text="No" AutoPostBack="True" OnCheckedChanged="raoStatic_CheckedChanged" /></td>
                </tr>
                <tr id="trOrientation" runat="server">
                    <td style="width: 25%">
                        <asp:Label ID="lblOrientation" runat="server" Text="Orientation:"></asp:Label></td>
                    <td style="width: 60%">
                        <asp:DropDownList ID="ddlOrientation" runat="server">
                        </asp:DropDownList></td>
                </tr>
                <tr id="trSignatureText" runat="server">
                    <td style="width: 25%;">
                        <asp:Label ID="lblSigText" runat="server" Text="Signaure Text:"></asp:Label></td>
                    <td style="width: 60%;">
                        <asp:ListBox ID="LstSignText" runat="server" Height="143px" SelectionMode="Multiple"
                            Width="260px"></asp:ListBox></td>
                </tr>
                <tr>
                    <td align="center" colspan="2" style="width: 25%;">
                        <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" Text="Submit" Width="80px" /></td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
