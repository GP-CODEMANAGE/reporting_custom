<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmAnzianoFilecopy.aspx.cs"
    Inherits="frmAnzanioFilecopy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Fund Admin WorkFlow</title>
    <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <link href="./common/CssClass.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <%--<script language="javascript" type="text/javascript">


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

            if (Sys.Browser.agent == Sys.Browser.InternetExplorer) {
                $get("<%=gvBilling.ClientID%>").tBodies[0].removeNode(true);
            }
            else {

                //            $get("<%=gvBilling.ClientID%>").innerHTML=””;
                document.getElementById("gvBilling").innerHTML = "";

            }
        }


    </script>--%>
    <script language="javascript" type="text/javascript">

        function onFolderPathClick(Path) {
            //$("#txtFolderPath").attr("value", Path);
            document.getElementById("TextBox1").value = Path;
        }
    </script>
    <%-- var windowObjectReference; // global variable

function openRequestedPopup() {
  windowObjectReference = window.open(
    Label2.text,
    "width=420,height=230,resizable,scrollbars=yes,status=1"
  );
}--%>
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
            text-align: right;
        }
        .Titlebig
        {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: normal;
            text-decoration: none;
        }
        
        .localLink
        {
            color: #1382CE;
            text-decoration: none;
            font-size: 10pt;
            border-spacing: 0 0;
            border-collapse: collapse;
        }
        .font
        {
            font-size: 10pt;
            
            
        }
         
    </style>
</head>
<body>
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
                    <br />
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
                <td class="Titlebig" colspan="3">
                    &nbsp;Fund Admin WorkFlow
                </td>
            </tr>
            <tr>
                <td valign="top" colspan="7" class="auto-style43">
                    <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                    <asp:Label ID="lblSavePopUp" runat="server"></asp:Label>
                    <br />
                    <asp:Label ID="lblMessageShow" runat="server" ForeColor="Red" Text="lblMessageShow"
                        Visible="False"></asp:Label>
                    <asp:Label ID="Label2" runat="server"></asp:Label>
                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ValidationGroup="vgBIE"
                        ShowMessageBox="true" DisplayMode="List" ShowSummary="false" />
                </td>
            </tr>
          <%--  <tr>
                <td class="style2">
                    <asp:Label ID="Label1" runat="server" Text="Email To:"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:Label ID="lblEmail" runat="server" Text=""></asp:Label>
                
                </td>
            </tr>--%>
            <tr>
                <td class="style2">
                    <asp:Label ID="Label3" runat="server" Text="Email CC(optional):"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:TextBox ID="txtMailBCC" runat="server" Width="364px"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlAdvisor" Display="Dynamic" ErrorMessage="Select Advisor" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%>
                    <br />
                </td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lblSelectAdv" runat="server" Text="Select Anziano folder"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <asp:DropDownList ID="ddlFolderName" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFolderName_SelectedIndexChanged"
                        TabIndex="1" Width="550px">
                    </asp:DropDownList>
                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlAdvisor" Display="Dynamic" ErrorMessage="Select Advisor" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                </td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lbltag" runat="server" Text="Select Tag"></asp:Label>
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlAdvisor" Display="Dynamic" ErrorMessage="Select Advisor" InitialValue="00000000-0000-0000-0000-000000000000" SetFocusOnError="True" ValidationGroup="vgBIE">*</asp:RequiredFieldValidator>--%><%-- <asp:ListBox ID="ListBox1" runat="server" Height="100%"></asp:ListBox>--%>
                    <asp:DropDownList ID="ddlPortalPath" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlPortalPath_SelectedIndexChanged"
                        TabIndex="1" Width="400px">
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:DropDownList ID="ddlYear" runat="server" AutoPostBack="True" TabIndex="1" 
                        Width="140px" >
                    </asp:DropDownList>
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                </td>
            </tr>
            <tr>
                <td class="style2">
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                    :
                </td>
                <td>
                </td>
            </tr>
        </table>
        <table>
            <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="false" Width="100%"
                 AlternatingRowStyle-Font-Size="Small" ControlStyle-CssClass="font" >
                <Columns>
                    <asp:BoundField DataField="Ssi_batchId" HeaderText="Ssi_batchId" Visible="False" />
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:CheckBox ID="chkbSelectBatch" runat="server" Checked='<%# Bind("FileNameMatchBool")%>'  Enabled='<%# Bind("HouseholdExists") %>' />
                             <%--<asp:CheckBox ID="chkbSelectBatch" runat="server" Checked='<%#  Bind("FileNameMatchBool") %>'  Enabled='<%# Eval("HouseholdExists").ToString() == "true" %> ' />--%>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <%-- <asp:BoundField DataField="FileName" HeaderText="File Name" SortExpression="BatchName"  />--%>
                    <asp:HyperLinkField DataNavigateUrlFields="PathUrl" DataTextField="FileName" Target="_blank"
                        HeaderText="File Name" ControlStyle-CssClass="localLink"  />
                    <asp:BoundField DataField="AnzianoID" HeaderText="AnzianoID" SortExpression="BatchName1"  />
                    <asp:BoundField DataField="LegalEntity" HeaderText="Legal Entity" SortExpression="LegalEntity"  />
                    <asp:BoundField DataField="HouseHold" HeaderText="HouseHold" SortExpression="HouseHold" />
                    <asp:TemplateField HeaderText="HouseHold" ControlStyle-CssClass="font">
                       <ItemTemplate>
                            <asp:DropDownList runat="server" ID="dtclients" DataSource='<%# newFolderstructure_client1() %>' DataValueField="Value" DataTextField="Text"
                        SelectedValue='<%# Bind("HouseHold") %>'/>
                            
                        </ItemTemplate>
                            
                        
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Verified">
                        <ItemTemplate>
                            <asp:CheckBox ID="chkVerify" runat="server" Checked='<%# Bind("Verified")%>'  />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle Height="10px" />
            </asp:GridView>
        </table>
        <br />
         <table>
            <asp:GridView ID="gvList1" runat="server" AutoGenerateColumns="false" 
                 AlternatingRowStyle-Font-Size="Small" ControlStyle-CssClass="font" >
                <Columns>
                    <asp:BoundField DataField="Ssi_batchId" HeaderText="Ssi_batchId" Visible="False" />
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:CheckBox ID="chkbSelectBatch" runat="server" Checked='<%# Bind("FileNameMatchBool") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <%-- <asp:BoundField DataField="FileName" HeaderText="File Name" SortExpression="BatchName"  />--%>
                    <asp:HyperLinkField DataNavigateUrlFields="PathUrl" DataTextField="FileName" Target="_blank"
                        HeaderText="File Name" ControlStyle-CssClass="localLink"  />
                   
                    <asp:TemplateField HeaderText="HouseHold" ControlStyle-CssClass="font">
                        <ItemTemplate>
                            <asp:DropDownList runat="server" ID="ddlClients1" DataSource='<%# newFolderstructure_client1() %>' DataValueField="Value" DataTextField="Text"
                        SelectedValue='<%# Bind("HouseHold") %>'>
                            </asp:DropDownList>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Verified">
                        <ItemTemplate>
                            <asp:CheckBox ID="chkVerify" runat="server" Checked='<%# Bind("Verified") %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle Height="10px" />
            </asp:GridView>
        </table>
        <%-- <asp:BoundField DataField="FileName" HeaderText="File Name" SortExpression="BatchName"  />--%>
        <br />
    </div>
    <asp:CheckBox ID="ChkTest" Checked="true" runat="server" />
    <asp:Label ID="lblTest" runat="server" Text="Test Mode"></asp:Label><br />
    <br />
    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" 
        Text="Start Fund Admin WorkFlow" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label>
    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
    </form>
    <%-- <asp:BoundField DataField="FileName" HeaderText="File Name" SortExpression="BatchName"  />--%><%-- <asp:BoundField DataField="FileName" HeaderText="File Name" SortExpression="BatchName"  />--%>