<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmcreateSLOA.aspx.cs" Inherits="frmcreateSLOA" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Create SLOA</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />

    <script src="./common/Calendar.js" type="text/javascript"></script>

    <style>
        .gvReportss {
            border-bottom: .02em solid #F2F2F2;
        }

        .ddcblkss {
            border-bottom: .01em solid #000000;
        }

        .gvReportssNo {
            border-bottom: .01em solid #ffffff;
        }

        .gvReportssBlack {
            border-bottom: .01em solid #000000;
        }

        .ddcblk {
            border-bottom: .02em solid #F2F2F2;
        }

        .ddcblksswhite {
            border-bottom: .01em solid #ffffff;
        }

        .BackgroundColor {
        }

        .familyname {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: bold;
            height: 25px;
        }

        .assetdistribution {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }

        .assDate {
            font-family: Frutiger 55 Roman;
            font-size: 10pt;
            font-style: italic;
        }

        .auto-style1 {
            height: 40px;
            width: 38%;
        }

        .auto-style2 {
            width: auto;
        }

        .auto-style4 {
            height: 40px;
            width: 22%;
        }
        .auto-style5 {
            width: auto;
        }
        .auto-style7 {
            height: 25px;
            width: 14%;
        }
       
        .auto-style8 {
            width:auto;
        }
        .auto-style9 {
            height: 25px;
            width: 9%;
        }
        .auto-style10 {
            width: auto;
        }
        .auto-style11 {
            height: 25px;
            width: auto;
        }
    </style>

    <script type="text/javascript">
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

        function ClearLabel() {
            document.getElementById("lblError").innerHTML = "";
        }



    </script>


</head>
<body>
    <form id="form1" runat="server">
        <table style="width: 100%">
            <%-- <a href="http://crm01/ISV/AdventReport/BatchReport/ReportReviewForm.aspx?id={Batch GUID(Batch)}">
            </a>--%>
            <tr>
                <td>
                    <table style="width: 100%">
                        <tr>
                            <td colspan="5">
                                <img src="images/Gresham_Logo__.jpg" />
                            </td>
                        </tr>
                    <%--    <tr>
                            <td colspan="5" class="Titlebig">Gresham Partners, LLC
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="Titlebig" colspan="5">
                                <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Large"
                                    Text="Create SLOA" Width="260px" Font-Names="Arial"></asp:Label></td>
                        </tr>
                        <tr>
                            <td style="height: 18px" valign="top" colspan="5">
                                <asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Names="Arial"></asp:Label>
                                  <asp:RequiredFieldValidator ID="RequiredFieldValidator1" ControlToValidate="lstHouseHold" runat="server" ErrorMessage="Please select household" ForeColor="Red" Font-Names="Arial"></asp:RequiredFieldValidator>
                            </td>

                        </tr>
                        <tr>
                            <td colspan="5" style="height: 18px" valign="top">
                                <asp:Label ID="lblMessage" runat="server" ForeColor="Red" Font-Names="Arial"></asp:Label></td>
                        </tr>


                        <tr>
                            <td class="auto-style10">
                                <asp:Label ID="lblHouseHold" runat="server" Text="Household:" Font-Names="Arial"></asp:Label></td>
                            <td class="auto-style4">
                                <asp:ListBox Font-Names="Arial" ID="lstHouseHold" runat="server" Height="220px"
                                    Width="220px"
                                    AutoPostBack="true" OnSelectedIndexChanged="lstHouseHold_SelectedIndexChanged1" onchange="ClearLabel();"></asp:ListBox>
                              
                            </td>
                            <td class="auto-style8">
                                <asp:Label ID="Label1" runat="server" Text="Legal Entity:" Font-Names="Arial"></asp:Label>
                            </td>

                            <td class="auto-style1">
                                <asp:ListBox ID="lstLegalEntity" runat="server" onchange="ClearLabel();" Height="150px" SelectionMode="Multiple" OnSelectedIndexChanged="lstLegalEntity_SelectedIndexChanged" Font-Names="Arial"></asp:ListBox></td>

                             <td style="width: 4px; height: 40px;"></td>
                        </tr>


                        <%--     <tr id="trLegalentity" runat="server">
                            <td style="width: 25%">
                                <br />
                                <asp:Label ID="Label4" runat="server" Text="Legal Entity:"></asp:Label></td>
                            <td class="auto-style1">
                                <br />
                                <asp:ListBox ID="lstLegalEntity" runat="server" onchange="ClearLabel();" Height="150px" SelectionMode="Multiple" OnSelectedIndexChanged="lstLegalEntity_SelectedIndexChanged"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px;"></td>
                        </tr>--%>

                        <tr>
                            <td style="font-family:Arial" align="left" class="auto-style10 ">Start Date
                                <br />
                                <br />
                                End Date
                            </td>
                            <td class="auto-style5">
                                <asp:TextBox ID="txtstartdate" runat="server" Font-Names="Arial"></asp:TextBox>
                                <a onclick="showCalendarControl(txtstartdate)">
                                    <img id="img1" alt="" border="0" src="images/calander.png" style="cursor: hand;" /></a><br />
                                <br />

                                <asp:TextBox ID="txtenddate" runat="server" Font-Names="Arial"></asp:TextBox>
                                <a onclick="showCalendarControl(txtenddate)">
                                    <img id="img2" alt="" border="0" src="images/calander.png" style="cursor: hand;" /></a>
                                
                            </td>

                            <td class="auto-style8">
                                <asp:Label ID="Label2" runat="server" Text="Fund:" Font-Names="Arial"></asp:Label></td>

                            <td class="auto-style2">
                                   <asp:ListBox ID="lstFund" runat="server" onchange="ClearLabel();" Height="150px" Width="300px" OnSelectedIndexChanged="lstFund_SelectedIndexChanged1" AutoPostBack="true" SelectionMode="Multiple" Font-Names="Arial"></asp:ListBox>
                            </td>
                             <td style="width: 4px; height: 40px;"></td>
                        </tr>

                        <%--    <tr>
                            <td style="width: 20%" align="left">End Date</td>
                            <td class="auto-style2">
                                <asp:TextBox ID="txtenddate" runat="server"></asp:TextBox>
                                <a onclick="showCalendarControl(txtenddate)">
                                    <img id="img2" alt="" border="0" src="images/calander.png" style="cursor: hand;" /></a>
                                <asp:CustomValidator ID="CustomValidator1" runat="server"
                                    ControlToValidate="txtenddate" Display="None" ErrorMessage="End date is not valid" ForeColor="Red"></asp:CustomValidator>
                                <asp:CompareValidator ID="CompareValidator1" runat="server" ErrorMessage="End date can not be less than start date"
                                    ControlToCompare="txtstartdate" ControlToValidate="txtenddate" Display="None"
                                    Operator="GreaterThanEqual" Type="Date" ForeColor="Red"></asp:CompareValidator></td>

                            <td style="width: 25%">
                                <asp:Label ID="Label2" runat="server" Text="Fund:"></asp:Label></td>
                            <td class="auto-style2">
                                <asp:ListBox ID="ListBox2" runat="server" onchange="ClearLabel();" Height="150px" Width="300px" OnSelectedIndexChanged="lstFund_SelectedIndexChanged1" AutoPostBack="true" SelectionMode="Multiple"></asp:ListBox>
                            </td>
                        </tr>--%>


                 <%--       <tr id="trFund" runat="server">
                            <td style="width: 25%">
                                <asp:Label ID="Label11" runat="server" Text="Fund:"></asp:Label></td>
                            <td class="auto-style1">
                                <asp:ListBox ID="lstFund" runat="server" onchange="ClearLabel();" Height="150px" Width="300px" OnSelectedIndexChanged="lstFund_SelectedIndexChanged1" AutoPostBack="true" SelectionMode="Multiple"></asp:ListBox></td>
                            <td style="width: 4px; height: 40px;"></td>
                        </tr>--%>


                        <tr>
                            <td class="auto-style11"></td>
                      <td class="auto-style7"></td><td class="auto-style9"></td><td class="auto-style7"></td><td class="auto-style7"></td>
                        </tr>


                        <tr>
                            <td align="left" class="auto-style10">&nbsp;</td>
                            <td class="auto-style5" >
                                <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click" Font-Names="Arial" />

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                                 

                            </td>

                             <td class="auto-style8">
                                 </td>

                              <td class="auto-style2">
                                    <asp:Button ID="Button2" runat="server" Text="Generate SLOA" OnClick="Button2_Click" Visible="false" Font-Names="Arial" />
                                  </td>


                        </tr>


                        <tr>
                            <td colspan="5" valign="top">
                                <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                    BackColor="White" BorderColor="#CCCCCC" OnRowDataBound="GridView1_RowDataBound"
                                    Font-Names="Arial" Font-Size="X-Small" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="2">
                                    <Columns>
                                        <asp:TemplateField HeaderStyle-HorizontalAlign="Left">
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="chkSelectNC" />
                                            </ItemTemplate>
                                            <HeaderTemplate>
                                                <input id="chkBoxAll" type="checkbox" onclick="checkAllBoxes()" />
                                            </HeaderTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField HeaderText="Household" DataField="Household" />
                                        <%--1--%>

                                        <asp:BoundField HeaderText="Legal Entity Name" DataField="Legal Entity Name" />
                                        <%--2--%>

                                        <asp:BoundField HeaderText="Fund Name" DataField="Fund Name" />
                                        <%--3--%>

                                        <asp:BoundField HeaderText="Contact Fulll Name" DataField="Contact Full Name" />
                                        <%--4--%>

                                        <asp:BoundField HeaderText="HouseholdId" DataField="HouseholdID" />
                                        <%--5--%>

                                        <asp:BoundField HeaderText="LegalEntityID" DataField="LegalEntityID" />
                                        <%--6--%>

                                        <asp:BoundField HeaderText="FundID" DataField="FundID" />
                                        <%--7--%>


                                        <asp:BoundField HeaderText="ContactId" DataField="ContactId" />
                                        <%--8--%>
                                    </Columns>
                                    <FooterStyle BackColor="White" ForeColor="#000066" />
                                    <RowStyle ForeColor="#000066" />
                                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                </asp:GridView>
                            </td>
                           
                        </tr>

                        <%--    <tr>
                            <td style="border-bottom: gray 1px solid; text-align: center; height: 12px;" colspan="2">
                                &nbsp;
                            </td>
                            <td style="width: 4px; height: 12px;">
                            </td>
                        </tr>--%>
                    </table>
                    <%-- <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                        ShowSummary="False" />--%>
                    &nbsp;
                    <input id="Hidden1" type="hidden" runat="Server" />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
