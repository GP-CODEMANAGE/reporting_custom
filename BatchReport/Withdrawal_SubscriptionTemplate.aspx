    <%@ Page Language="C#" AutoEventWireup="true" CodeFile="Withdrawal_SubscriptionTemplate.aspx.cs" Inherits="BatchReport_Withdrawal_SubscriptionTemplate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
   <%-- <title>Withdrawal And Subscription Template</title>--%>
     <title>Create Additional Subscription and Withdrawal Letters</title>
    
    <script language="Javascript" type="text/javascript">
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
    </script>
    <style type="text/css">
        .auto-style1 {
            width: 80%;
        }
        .auto-style2 {
            height: 23px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 100%">
                <tr>
                    <td colspan="2">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="Titlebig">Gresham Partners, LLC
                    </td>
                </tr>
                <tr>
                    <td class="Titlebig" colspan="4">
                        <asp:Label ID="lblFilterHeader" runat="server" Font-Bold="True" Font-Size="Large"
                            Text="Create Additional Subscription and Withdrawal Letters" Width="389px"></asp:Label>

                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td class="auto-style1"></td>
                    
                </tr>
                <tr>
                    <td colspan="3" class="auto-style2">
                        <asp:Label ID="lblError" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                        
                    </td>
                </tr>
                <tr>
                    <td colspan="3" class="auto-style2">
                           <asp:Label ID="lblSuccess" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                        
                    </td>
                </tr>

               


                <tr>
                    <td style="width: 20%">Associate:
                    </td>
                    <td class="auto-style1">&nbsp;<asp:DropDownList ID="ddlAssociate" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAssociate_SelectedIndexChanged">
                    </asp:DropDownList>
                    </td>
                   
                </tr>
                <tr>
                    <td style="width: 20%">Household:
                    </td>
                    <td class="auto-style1">&nbsp;<asp:DropDownList ID="ddlHouseHold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHouseHold_SelectedIndexChanged" Style="height: 22px">
                    </asp:DropDownList>
                    </td>
                   
                </tr>
                <tr>
                    <td style="width: 20%">Legal Entity:
                    </td>
                    <td class="auto-style1">&nbsp;<asp:DropDownList ID="ddlLegalEntity" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlLegalEntity_SelectedIndexChanged">
                    </asp:DropDownList>
                    </td>
                    
                </tr>
                   <tr>
                    <td style="width: 20%">Recommendation Status:
                    </td>
                    <td class="auto-style1">&nbsp;<asp:DropDownList ID="ddlStatus" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlStatus_SelectedIndexChanged">
                    </asp:DropDownList>
                    </td>
                    
                </tr>
                <tr>
                    <td >
                        </td>
                      <td >

                             &nbsp;<asp:Button Font-Names="Verdana" ID="btnSubmit" runat="server" Text="Create Letters" OnClick="btnSubmit_Click" style="height: 26px" />

                        </td>
                    
                    </tr>
                 <tr>
                    <td align="left" > <asp:Label ID="lblCount"  runat="server" Font-Bold="True" ></asp:Label>
                        </td>
                    <td  >
                       
                        </td>
                    
                    </tr>
                <tr>
                    <td colspan ="2">
                        <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                                   BackColor="White" BorderColor="#CCCCCC"
                                    Font-Names="Verdana"  Font-Size="X-Small" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="2">
                                    <Columns>
                                        <asp:TemplateField Visible ="false">
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="chkSelectNC" />
                                            </ItemTemplate>
                                            <HeaderTemplate>
                                                <input id="chkBoxAll" type="checkbox" onclick="checkAllBoxes()" />
                                            </HeaderTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField HeaderText="Household" DataField="Household" />
                                        <asp:BoundField HeaderText="Legal Entity" DataField="Legal Entity" />
                                        <asp:BoundField HeaderText="Close Date" DataField="Close Date" />
                                        <asp:BoundField HeaderText="Fund" DataField="Fund" />
                                        <asp:BoundField HeaderText="Transaction Type" DataField="Transaction Type" />
                                        <asp:BoundField HeaderText="Withdrawal Type" DataField="Withdrawal type" />
                                        <asp:BoundField HeaderText="Percentage" DataField="Percentage" DataFormatString="{0:P1}"/>                                                                            
                                        <asp:BoundField HeaderText="Confirmed Amount" DataField="Confirmed Amount"  DataFormatString="{0:C0}" Visible="True" />
                                        <asp:BoundField HeaderText="Ssi_transactionrecommendationId" DataField="Ssi_transactionrecommendationId" Visible="False" />
                                       <%--<asp:BoundField HeaderText="FileName" DataField="FileName" visible="false"/>--%>
                                          <asp:BoundField HeaderText="Amount" DataField="Amount" visible="false"/>
                                         <asp:BoundField HeaderText="Year" DataField="YearNmb"  visible="false"/>
                                         <asp:BoundField HeaderText="GRASSeries" DataField="ListTxtGRAS"  visible="false"/>
                                        <asp:BoundField HeaderText="GPESSeries" DataField="ListTxtGPES"  visible="false"/>
                                          <%--<asp:BoundField HeaderText="Signer1Name" DataField="Signer1Name"  visible="false"/>--%><%--11--%>
                                          <%--<asp:BoundField HeaderText="Signer1Title" DataField="Signer1Title"  visible="false"/>--%><%--12--%>
                                         <%--<asp:BoundField HeaderText="Signer2Name" DataField="Signer2Name" visible="false" />--%><%--13--%>
                                         <%-- <asp:BoundField HeaderText="Signer2Title" DataField="Signer2Title"  visible="false"/>--%><%--14--%>
                                            <%--<asp:BoundField HeaderText="Signer3Name" DataField="Signer3Name"  visible="false"/>--%><%--15--%>
                                          <%--<asp:BoundField HeaderText="Signer3Title" DataField="Signer3Title"  visible="false"/>--%><%--16--%>

                                        <%--<asp:BoundField HeaderText="LE_Signer1Name" DataField="LE_Signer1Name"  visible="false"/>--%><%--17--%>
                                          <%--<asp:BoundField HeaderText="LE_Signer1Title" DataField="LE_Signer1Title"  visible="false"/>--%><%--18--%>
                                         <%--<asp:BoundField HeaderText="LE_Signer2Name" DataField="LE_Signer2Name" visible="false" />--%><%--19--%>
                                          <%--<asp:BoundField HeaderText="LE_Signer2Title" DataField="LE_Signer2Title"  visible="false"/>--%><%--20--%>
                                            <%--<asp:BoundField HeaderText="LE_Signer3Name" DataField="LE_Signer3Name"  visible="false"/>--%><%--21--%>
                                          <%--<asp:BoundField HeaderText="LE_Signer3Title" DataField="LE_Signer3Title"  visible="false"/>--%><%--22--%>

                                         <%--<asp:BoundField HeaderText="ShowAccSignorFlg" DataField="ShowAccSignorFlg"  visible="false"/>--%><%--23--%>

                                    </Columns>
                                    <FooterStyle BackColor="White" ForeColor="#000066" />
                                    <RowStyle ForeColor="#000066" />
                                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                </asp:GridView>
                        </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
