<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BillingExcludeAssets.aspx.cs"
    Inherits="BillingExcludeAssets" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Billing Worksheet Form</title>
    <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript">


        function LOAD() {
            Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(clearDisposableItems);
        }

        function Confirm() {

            var confirm_value = document.createElement("INPUT");
            confirm_value.type = "hidden";
            confirm_value.name = "confirm_value";
            if (confirm("Do you want to save data?")) {
                confirm_value.value = "Yes";
            } else {
                confirm_value.value = "No";
            }
            document.forms[0].appendChild(confirm_value);
        }
        function onlyNumbers(evt) {
            var e = event || evt; // for trans-browser compatibility
            var charCode = e.which || e.keyCode;
            var result = null;
            if (charCode > 31 && (charCode < 48 || charCode > 57))
                result = false;
            else
                result = true;


            if (charCode == 46 || charCode == 45)
                result = true;

            return result;
        }

        function clearDisposableItems(sender, args) {

        if (Sys.Browser.agent == Sys.Browser.InternetExplorer ) {
         $get("<%=gvBilling.ClientID%>").tBodies[0].removeNode(true);
         }
         else {

//            $get("<%=gvBilling.ClientID%>").innerHTML=””;
            document.getElementById("gvBilling").innerHTML = "";

        }
        }


    </script>





    <style type="text/css">
        .style1
        {
            width: 12px;
        }
        .auto-style34
        {
            width: 1263px;
        }
        .style2
        {
            width: 187px;
        }
        .style3
        {
            width: 83px;
        }
        .TextAlgRgh
        {
            text-align  :right;
        }
        
    </style>
</head>
<body >
    <form id="form1" runat="server">
    <div onload="LOAD()">
     
        <table class="auto-style34">
            <tr>
                <td colspan="3" class="auto-style35">
                    <img src="images/Gresham_Logo__.jpg" />
                </td>
                <td class="style1">
                    &nbsp;
                </td>
                <td class="style3">
                    &nbsp;
                </td>
                <td class="auto-style37">
                    &nbsp;
                </td>
                <td class="auto-style38">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="3" class="auto-style39">
                    Gresham Partners, LLC
                </td>
                <td class="style1">
                    &nbsp;
                </td>
                <td class="style3">
                    &nbsp;
                </td>
                <td class="auto-style41">
                    &nbsp;
                </td>
                <td class="auto-style42">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="auto-style39" colspan="3">
                    Billing Worksheet Form
                </td>
                <td class="style1">
                    &nbsp;
                </td>
                <td class="style3">
                    &nbsp;
                </td>
                <td class="auto-style41">
                    &nbsp;
                </td>
                <td class="auto-style42">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td valign="top" colspan="7" class="auto-style43">
                    <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                    <asp:Label ID="lblSavePopUp" runat="server"></asp:Label>
                    <br />
                    <asp:Label ID="lblMessageShow" runat="server" ForeColor="Red" 
                        Text="lblMessageShow" Visible="False"></asp:Label>
                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ValidationGroup="vgBIE"
                        ShowMessageBox="true" DisplayMode="List" ShowSummary="false" />
                </td>
            </tr>
           <tr>
                <td class="style2">
                    <asp:Label ID="Label1" runat="server" Text="AUM as of Date"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                </td>
                <td>
                    <asp:TextBox ID="txtAUMDate" runat="server" Width="119px" AutoPostBack=true  TabIndex="4" 
                        ontextchanged="txtAUMDate_TextChanged"></asp:TextBox>
                    <a onclick="showCalendarControl(txtAUMDate)">
                        <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator
                            ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtAUMDate" Display="Dynamic"
                            ErrorMessage="Select AUM as of Date" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
           
                <td class="style2">
                    <asp:Label ID="lblSelectAdv" runat="server" Text="Advisor"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:DropDownList ID="ddlAdvisor" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlAdvisor_SelectedIndexChanged"
                        TabIndex="1"  Width="221px" >
                    </asp:DropDownList>
                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlAdvisor" Display="Dynamic" ErrorMessage="Select Advisor" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%>
                </td>
                <%-- <td class="auto-style47" style="width: 30px">&nbsp;</td>
                    <td class="auto-style49">&nbsp;</td>
                    <td class="auto-style50">&nbsp;</td>
                    <td class="auto-style51">&nbsp;</td>--%>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lblHH" runat="server" Text="Household"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:DropDownList ID="ddlHH" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHH_SelectedIndexChanged"
                        TabIndex="2" Width="221px">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="ddlHH"
                        Display="Dynamic" ErrorMessage="Select Household" InitialValue="00000000-0000-0000-0000-000000000000"
                        SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                </td>
                <%-- <td valign="top" class="auto-style47">&nbsp;</td>
                    <td valign="top" class="auto-style49">&nbsp;</td>

                    <td class="auto-style51">&nbsp;</td>--%>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lblBilFor" runat="server" Text="Billing For"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:DropDownList ID="ddlBillFor" runat="server" AutoPostBack="True" TabIndex="3"
                        Width="221px" onselectedindexchanged="ddlBillFor_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlBillFor"
                        Display="Dynamic" ErrorMessage="Select Billing For" InitialValue="ALL" SetFocusOnError="True"
                        ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                </td>
                <%-- <td valign="top" class="auto-style47">&nbsp;</td>
                    <td valign="top" class="auto-style49">&nbsp;</td>
                    <td valign="top" class="auto-style50">&nbsp;</td>--%><%-- <td class="auto-style51">
                        <a onclick="showCalendarControl(txtBillingPeriod)">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="txtBillingPeriod" Display="Dynamic" ErrorMessage="Select Billing period" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>
                    </td>--%>
            </tr>
              

           
            <tr>
                <td class="style2">
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                </td>
                <td style="text-align:justify">
                    <asp:Button ID="btnSubmit" runat="server" Text="Submit" 
                        OnClick="btnSubmit_Click" Visible="False" />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                           
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                       </td>
            </tr>
            <tr>
                <td class="style2">
                </td>
               <td class="auto-style46">
                </td>
                 <td>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;


        <asp:Label ID="lblBillingFor" runat="server" Text="lblBillingFor" Font-Bold="True" 
                         Font-Size="X-Large" Visible="False"></asp:Label>

                     <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <asp:Label ID="lblAUMDate" runat="server" Font-Bold="True" Font-Size="X-Large" 
                         Text="lblAUMDate" Visible="False"></asp:Label>
                     <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <asp:Button ID="btnSave" runat="server" Text="Submit" onclick="btnSave_Click"  
                        Visible=false Width="85px" />
                      
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                       <asp:Button ID="btnSaveGenrate" runat="server" Text="Submit and Generate"   
                        Visible=false Width="163px" onclick="btnSaveGenrate_Click" Height="28px" />
                     <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                </td>
                 <td class="style1">
                </td>
              
                </td>
                 <td class="auto-style46">
                     &nbsp;</td>
                 
             </tr>
        </table>
        <br />
      <asp:ScriptManager ID="ScriptManager1" runat="server" >
</asp:ScriptManager>
       <asp:UpdatePanel ID="update1" runat="server">
    <ContentTemplate>
        <asp:GridView ID="gvBilling" runat="server" Width="1700px" 
            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="10"
            Style="margin-top: 0px"   >
            <Columns>
                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                <asp:BoundField DataField="PosAssetClassName" HeaderText="Asset Class" />    <%--0--%> 
                <asp:BoundField DataField="Reporting Name" HeaderText="Reporting Name" />    <%--1--%> 
               <asp:BoundField DataField="Meeting book tab2 name" HeaderText="Column Name" />       <%--2--%> 
                <asp:BoundField DataField="AccountName" HeaderText="Legal Entity" />               <%--2--%>   <%--3--%> 
                <%--<asp:BoundField DataField="SecurityName" HeaderText="Security Name" />--%>
                <%--<asp:BoundField DataField="ssi_MarketValue" HeaderText="Market Value" />--%>
                <asp:BoundField DataField="ssi_BillingMarketValue" HeaderText="Actual" DataFormatString="{0:C2}"  ItemStyle-HorizontalAlign="Right"/>     <%--3--%>   <%--4--%> 
                <%--   <asp:BoundField DataField="FinalBillingMarketValue" HeaderText="Billing" />--%>
                
                <asp:TemplateField HeaderText="Billing" ControlStyle-Width="120px">
                    <ItemTemplate>
                        <asp:TextBox ID="txtBilling" runat="server" Width="100px" AutoPostBack="true" onkeypress="return onlyNumbers(this);" Font-Size="10" CssClass="TextAlgRgh"
                            OnTextChanged="txtBilling_TextChanged" Text='<%# Bind("FinalBillingMarketValue") %>' DataFormatString="{0:C2}"></asp:TextBox>
                        <%--  <asp:RegularExpressionValidator ID="regUnitsInStock" runat="server" ControlToValidate="txtBilling" ErrorMessage="Only numbers allowed" AutoPostBack="true" 
                            ValidationExpression="(^([0-9]*\d*\d{1}?\d*)$)" Display="Dynamic"></asp:RegularExpressionValidator>Onkeypress="return onlyNumbers(this);--%>
                    </ItemTemplate>
                </asp:TemplateField>    <%--4--%>   <%--5--%>  
                <asp:TemplateField HeaderText="Exclude Billing" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:CheckBox runat="server" ID="cbExcludeBilling" AutoPostBack="true" OnCheckedChanged="cbExcludeBilling_CheckedChanged"
                            Checked='<%# bool.Parse(Eval("BillingExcludeFlg").ToString()) %>' Enable='<%# !bool.Parse(Eval("BillingExcludeFlg").ToString()) %>'>
                        </asp:CheckBox>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>  <%--5--%>   <%--6--%>  
                <%--    <asp:BoundField DataField="FinalAUMMarketValue" HeaderText="AUM" />--%>
                <asp:TemplateField HeaderText="AUM" ControlStyle-Width="120px">
                    <ItemTemplate>
                        <asp:TextBox ID="txtAUM" runat="server" Width="100px" onkeypress="return onlyNumbers(this);" Font-Size="10" CssClass="TextAlgRgh"
                            Text='<%# Bind("FinalAUMMarketValue") %>' OnTextChanged="txtAUM_TextChanged"
                            AutoPostBack="true"></asp:TextBox>
                        <%--   <asp:CompareValidator ID="CompareValidator2"  runat="server" Operator="DataTypeCheck" Type="Integer" ControlToValidate="txtAUM" ErrorMessage="Number Only" />--%>
                    </ItemTemplate>
                </asp:TemplateField>  <%--6--%>   <%--7--%>  
                <asp:TemplateField HeaderText="Exclude AUM" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:CheckBox runat="server" ID="cbExcludeAum" AutoPostBack="true" OnCheckedChanged="cbExcludeAUM_CheckedChanged"
                            Checked='<%# bool.Parse(Eval("AUMExcludeFlg").ToString()) %>' Enable='<%# !bool.Parse(Eval("AUMExcludeFlg").ToString()) %>'>
                        </asp:CheckBox>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>  <%--7--%>    <%--8--%>  
                <asp:TemplateField HeaderText="Billing Fee Percentage">
                    <ItemTemplate>
                        <asp:TextBox ID="txtBillPer" runat="server" Width="50px" onkeypress="return onlyNumbers(this);" Font-Size="10" OnTextChanged="txtBillPer_TextChanged" CssClass="TextAlgRgh"
                            Text='<%# Bind("BillingFeePct") %>' AutoPostBack="true"></asp:TextBox>% 
                        <%--AutoPostBack="True" OnTextChanged="txtComment_TextChanged" OnTextChanged="txtBillPer_TextChanged" --%>
                    </ItemTemplate>
                </asp:TemplateField>   <%--8--%>   <%--9--%>  
                <asp:BoundField DataField="ssi_accountid" HeaderText="AccountID" />
                <%--9--%>  <%--10--%>  
                <asp:BoundField DataField="ssi_securityid" HeaderText="SecurityID" />
                <%--10--%>  <%--11--%> 
                <asp:BoundField DataField="sas_assetclassid" HeaderText="Asset Class ID" />
                <%--11--%> <%--12--%> 
                <asp:BoundField DataField="Ssi_billingId" HeaderText="Billing ID" Visible="false" />
                <%--12--%> <%--13--%> 
                <asp:BoundField DataField="ACLevelBillingExceptionFlg" HeaderText="ACLevelBillingExceptionFlg"
                    Visible="false" />
                <%--13--%>  <%--14--%> 
                <asp:BoundField DataField="ACLevelAUMExceptionFlg" HeaderText="ACLevelAUMExceptionFlg"
                    Visible="false" />
                <%--14--%>  <%--15--%> 
                <asp:BoundField DataField="AssetLevelFlg" HeaderText="AssetLevelFlg" Visible="false" />
                <%--15--%>  <%--16--%>  
                <asp:BoundField DataField="TotalLevelFlg" HeaderText="TotalLevelFlg" Visible="false" />
                <%--16--%>  <%--17--%> 
                <asp:BoundField DataField="BillingExceptionId" HeaderText="BillingExceptionId" Visible="false" />
                <%--17--%>  <%--18--%> 
                <asp:BoundField DataField="BillingFeeExceptionId" HeaderText="BillingFeeExceptionId"
                    Visible="false" />
                <%--18--%>  <%--19--%> 
                <asp:BoundField DataField="BillingExceptionType" HeaderText="BillingExceptionType"
                    Visible="false" />
                <%--19--%>  <%--20--%> 
                <asp:BoundField DataField="AUMExceptionType" HeaderText="AUMExceptionType" Visible="false" />
                <%--20--%>  <%--21--%> 
                <asp:BoundField DataField="IdNmb" HeaderText="IdNmb" />
                <%--21--%>  <%--22--%> 
                <asp:BoundField DataField="BillingExcludeFlg" HeaderText="BillingExcludeFlg" />
                <%--22--%>  <%--23--%> 
                <asp:BoundField DataField="AUMExcludeFlg" HeaderText="AUMExcludeFlg" />
                <%--23--%>  <%--24--%> 
                <asp:BoundField DataField="BillingExtra" HeaderText="BillingExtra" />
                <%--24--%>  <%--25--%> 
                <asp:BoundField DataField="AUMExtra" HeaderText="AUMExtra" />
                <%--25--%>  <%--26--%> 
                <asp:BoundField DataField="BillingFeePct" HeaderText="BillingFeePct" />
                <%--26--%>  <%--27--%> 
                <asp:BoundField DataField="BillingFeePctExtra" HeaderText="BillingFeePctExtra" />
                <%--27--%>  <%--28--%> 
                <asp:BoundField DataField="FinalBillingMarketValue" HeaderText="FinalBillingMarketValue" />
                <%--28--%> <%--29--%> 
                 <asp:BoundField DataField="FinalAUMMarketValue" HeaderText="FinalAUMMarketValue" />
                <%--29--%>  <%--30--%> 
                 <asp:BoundField DataField="BillingFeePct" HeaderText="BillingFeePct" />
                <%--30--%>  <%--31--%> 
                 <asp:BoundField  HeaderText="SubtractBilling" />
                <%--31--%>  <%--32--%> 
                 <asp:BoundField  HeaderText="SubtractAum" />
                <%--32--%>  <%--33--%> 
                <asp:BoundField DataField="ssi_Startdate" HeaderText="ssi_Startdate" />
                <%--33--%>  <%--34--%> 
                <asp:BoundField DataField="ssi_EndDate" HeaderText="ssi_EndDate" />
                <%--34--%>  <%--35--%> 
                <asp:BoundField DataField="BillingAumExceptionID" HeaderText="BillingAumExceptionID" />
                <%--35--%>  <%--36--%> 
                <asp:BoundField DataField="BillingAumFeeExceptionID" HeaderText="BillingAumFeeExceptionID" />
                <%--36--%>  <%--37--%> 
                <asp:BoundField DataField="ssi_AUMStartDate" HeaderText="ssi_AUMStartDate" />
                <%--37--%>  <%--38--%> 
                <asp:BoundField DataField="ssi_AUMEndDate" HeaderText="ssi_AUMEndDate" />
                <%--38--%> <%--39--%> 
                <asp:BoundField DataField="MinAUMDate" HeaderText="MinAUMDate" />
                <%--39--%>  <%--40--%> 
                <asp:BoundField DataField="MinDate" HeaderText="MinBillingDate" />
                <%--40--%><%--41--%> 
                <asp:BoundField  HeaderText="BillingFlags" />
                <%--41--%> <%--42--%> 
                <asp:BoundField  HeaderText="AumFlags" />
                <%--42--%> <%--43--%> 
                <asp:BoundField  HeaderText="PositiveBilling" />
                <%--43--%> <%--44--%> 
                <asp:BoundField  HeaderText="SecLEvel" />
                <%--44--%> <%--45--%> 
                <asp:BoundField  HeaderText="PositiveAUM" />   
                <%--45--%> <%--46--%> 
                <asp:BoundField  HeaderText="MinFutureDate" DataField="MinFutureDate" />   
                <%--47--%> 
                  <asp:BoundField  HeaderText="MinAumFutureDate" DataField="MinAumFutureDate" />   
               <%--48--%> 
                <asp:BoundField  HeaderText="ColourBillingExceptionType" DataField="ColourBillingExceptionType" />   
               <%--49--%> 
               <asp:BoundField  HeaderText="ColourAUMExceptionType" DataField="colourAUMExceptionType" />   
               <%--50--%> 
               <asp:BoundField  HeaderText="DiffFlg" DataField="DiffFlg" />
                 <%--51--%> 
                   <asp:TemplateField HeaderText="AA" ControlStyle-Width="120px">
                    <ItemTemplate>
                        <asp:TextBox ID="txtAA" runat="server" Width="100px" onkeypress="return onlyNumbers(this);" Font-Size="10" CssClass="TextAlgRgh"
                            Text='<%# Bind("FinalAAMarketValue") %>' OnTextChanged="txtAA_TextChanged"
                            AutoPostBack="true"></asp:TextBox>                       
                    </ItemTemplate>
                </asp:TemplateField> <%--52--%> 
                <asp:TemplateField HeaderText="Exclude AA" ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:CheckBox runat="server" ID="cbExcludeAA" AutoPostBack="true" OnCheckedChanged="cbExcludeAA_CheckedChanged"
                            Checked='<%# bool.Parse(Eval("AAExcludeFlg").ToString()) %>' Enable='<%# !bool.Parse(Eval("AAExcludeFlg").ToString()) %>'>
                        </asp:CheckBox>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField> <%--53--%> 
                  <asp:BoundField DataField="AAExcludeFlg" HeaderText="AAExcludeFlg" Visible="false" />
                <%--54--%> 
                <asp:BoundField DataField="ACLevelAAExceptionFlg" HeaderText="ACLevelAAExceptionFlg"
                    Visible="false" />
                <%--55--%> 
                   <asp:BoundField  HeaderText="AAFlags" Visible="false"/>
                   <%--56--%> 
                      <asp:BoundField  HeaderText="ColourAAExceptionType" DataField="colourAAExceptionType" Visible="false" />  
                <%--57--%> 
                 <asp:BoundField  HeaderText="PositiveAA" Visible="false"/>  
                        <%--58--%> 
                <asp:BoundField  HeaderText="SubtractAA" Visible="false"/>
                 <%--59--%> 
                   <asp:BoundField DataField="AAExtra" HeaderText="AAExtra" Visible="false" />
                 <%--60--%> 
                 <asp:BoundField DataField="AAExceptionType" HeaderText="AAExceptionType" Visible="false" />
               <%--61--%> 
                 <asp:BoundField DataField="BillingAAExceptionID" HeaderText="BillingAAExceptionID" Visible="false"/>
                <%--62--%> 
                <asp:BoundField DataField="BillingAAFeeExceptionID" HeaderText="BillingAAFeeExceptionID" Visible="false"/>
               <%--63--%> 
                 <asp:BoundField  HeaderText="MinAAFutureDate" DataField="MinAAFutureDate" Visible="false"/>   
               <%--64--%> 
                 <asp:BoundField DataField="FinalAUMMarketValue" HeaderText="FinalAUMMarketValue" Visible="false"/>
                <%--65--%> 
                  <asp:BoundField DataField="ssi_AAStartDate" HeaderText="ssi_AAStartDate" Visible="false"/>
                <%--66--%>  
                <asp:BoundField DataField="MinAADate" HeaderText="MinAADate" Visible="false"/>
                <%--67--%> 
                 <asp:TemplateField HeaderText="AA Category">
                                            <ItemTemplate>
                                                <asp:DropDownList runat="server" ID="ddlAACategory" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                           <%--68--%> 
                 <asp:BoundField DataField="ssi_AACategory" HeaderText="ssi_AACategory" Visible="False" />
                <%--69--%> 











                <%-- <asp:TemplateField>
                    <ItemTemplate>
                        <%#Container.DataItemIndex+1 %>
                    </ItemTemplate>
                </asp:TemplateField>--%>
            </Columns>
            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
            <RowStyle ForeColor="#000066" />
             <HeaderStyle Height="10px"  Font-Size="10" />
        </asp:GridView>

       </ContentTemplate>
       </asp:UpdatePanel>
    </div>
    <p style="margin-left: 840px">
                     <asp:Button ID="btnSave0" runat="server" Text="Submit" onclick="btnSave_Click"  
                        Visible=false Width="85px" />
                      
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                       <asp:Button ID="btnSaveGenrate0" runat="server" Text="Submit and Generate"   
                        Visible=false Width="163px" onclick="btnSaveGenrate_Click" />
                     </p>
    </form>
</body>
</html>
