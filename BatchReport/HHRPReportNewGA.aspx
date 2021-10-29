<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HHRPReportNewGA.aspx.cs"
    Inherits="BatchReport_HHRPReportNewGA" Culture="en-US" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Advent Individual Reports GA 2.0</title>
    <link href="../common/Calendar.css" rel="stylesheet" type="text/css" />
    <link id="style1" href="../common/gresham.css" rel="stylesheet" type="text/css" />

    <script src="../common/Calendar.js" type="text/javascript"></script>

    <script language="javascript" type="text/javascript">

        function ClearValues() {
            //debugger;//
            var ddlHHRP = document.getElementById("ddlHHRP").value;
            var ddlReportType = document.getElementById("ddlReportType").value;
            var txtPriorDate = document.getElementById("txtPriorDate").value;
            var txtEndDate = document.getElementById("txtEndDate").value;

            if (ddlHHRP != "0" && ddlReportType != "0" && txtPriorDate != "" || txtEndDate != "" || document.getElementById("chkNoComparison").checked == false || document.getElementById("chkNoComparison").checked == true || document.getElementById("chkConvertToAssetDistComp").checked == false || document.getElementById("chkConvertToAssetDistComp").checked == true || document.getElementById("chkSuppressManagerDetail").checked == false || document.getElementById("chkSuppressManagerDetail").checked == true) {
                lblError.innerText = "";
                return false;
            }

        }

         
        function ValidateDate(source, args) {
            
            var txtPriorDate = document.getElementById("txtPriorDate").value;
            var txtEndDate = document.getElementById("txtEndDate").value;

            if (txtPriorDate != "") {
            var txtDate = new Date(txtPriorDate);
            var txtToDate = new Date(txtEndDate);           
           
                if (txtDate <= txtToDate)
                    args.IsValid = true;
                else
                    args.IsValid = false;
            }
        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table style="width: 839px">
                <tr>
                    <td colspan="3">
                        <img src="images/Gresham_Logo__.jpg" /></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan="3" class="Titlebig">Gresham Partners, LLC
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan="3" class="Titlebig">
                        <asp:Label ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large" Text="Advent Individual Reports GA 2.0"></asp:Label>
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 28px">
                        <asp:Label ID="lblError" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan="3" style="height: 18px">
                        <table style="height: 158px" width="100%">
                              <tr>
                                <td style="width: 281px">Household Type :
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlHouseHoldType" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHouseHoldType_SelectedIndexChanged" >
                                    </asp:DropDownList>
                                </td>
                                <td>
                                    <%--<asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="ddlHouseHoldType"
                                        Display="Dynamic" ErrorMessage="Please select Household Type" InitialValue="0">*</asp:RequiredFieldValidator>--%></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td style="width: 281px">Household :
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlHouseHold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlHouseHold_SelectedIndexChanged">
                                    </asp:DropDownList>
                                </td>
                                <td>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="ddlHouseHold"
                                        Display="Dynamic" ErrorMessage="Please select Household" InitialValue="0">*</asp:RequiredFieldValidator></td>
                                <td></td>
                            </tr>
                            <tr>
                                <td style="width: 281px">Report Type :
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlReportType" runat="server" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlReportType_SelectedIndexChanged" onchange="ClearValues();">
                                        <asp:ListItem Value="0">Please Select</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                                <td>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlHouseHold"
                                        Display="Dynamic" ErrorMessage="Please select Household" InitialValue="0">*</asp:RequiredFieldValidator>
                                </td>
                                <td></td>
                            </tr>
                            <tr>
                                <td style="width: 281px">Household Report Parameter : &nbsp;
                                </td>
                                <td colspan="3">
                                    <asp:DropDownList ID="ddlHHRP" runat="server" onchange="ClearValues();" >
                                        <asp:ListItem Value="0">Please Select</asp:ListItem>
                                    </asp:DropDownList><asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server"
                                        ControlToValidate="ddlHHRP" Display="Dynamic" ErrorMessage="Please select Household Report Parameter"
                                        InitialValue="0">*</asp:RequiredFieldValidator></td>
                            </tr>
                            <tr id="Tr1" runat="server">
                                <td style="width: 281px">Prior Date :</td>
                                <td>
                                    <asp:TextBox ID="txtPriorDate" runat="server" onchange="ClearValues();"></asp:TextBox>&nbsp;
                                    <a onclick="showCalendarControl(txtPriorDate)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" runat="server" />&nbsp;&nbsp;</a><asp:RegularExpressionValidator
                                            ID="RegularExpressionValidator2" runat="server" ErrorMessage="Invalid Prior Date"
                                            ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
                                            ControlToValidate="txtPriorDate" Display="Dynamic">*</asp:RegularExpressionValidator>
                                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
                                    <asp:CheckBox ID="chkNoComparison" runat="server" Text="no comparison line" onclick="ClearValues();"
                                        AutoPostBack="True" OnCheckedChanged="chkNoComparison_CheckedChanged" Font-Bold="True" /></td>
                                <td>&nbsp;
                                </td>
                                <td>&nbsp;</td>
                            </tr>
                            <tr id="Tr2" runat="server">
                                <td style="height: 26px; width: 281px;">As Of &nbsp;Date :</td>
                                <td style="height: 26px">
                                    <asp:TextBox ID="txtEndDate" runat="server" onchange="ClearValues();"></asp:TextBox>&nbsp;<a
                                        onclick="showCalendarControl(txtEndDate)">
                                        <img id="Img2" alt="" border="0" src="images/calander.png" /></a>
                                </td>
                                <td style="height: 26px">
                                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="Invalid End Date"
                                        ControlToValidate="txtEndDate" ValidationExpression="^(?:(?:(?:0?[13578]|1[02])(\/|-|)31)\1|(?:(?:0?[13-9]|1[0-2])(\/|-|)(?:29|30)\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:0?2(\/|-|)29\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:(?:0?[1-9])|(?:1[0-2]))(\/|-|)(?:0?[1-9]|1\d|2[0-8])\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$"
                                        Display="Dynamic">*</asp:RegularExpressionValidator>
                                <%--    <asp:CompareValidator ID="CompareValidator1" runat="server" ControlToCompare="txtPriorDate"
                                        ControlToValidate="txtEndDate" ErrorMessage="Please select greater date than Prior Date"
                                        Operator="GreaterThan" Type="Date">*</asp:CompareValidator>--%>
                                    <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="Please select greater date than Prior Date"
                                        ClientValidationFunction="ValidateDate" >*</asp:CustomValidator>

                                </td>
                                <td style="height: 26px"></td>
                            </tr>
                            <tr id="Tr3" runat="server">
                                <td style="width: 281px; height: 26px">
                                    <span style="font-size: 7pt"></span>Convert to Asset Distribution Comparison :</td>
                                <td style="height: 26px">
                                    <asp:CheckBox ID="chkConvertToAssetDistComp" runat="server" onclick="ClearValues();" /></td>
                                <td style="height: 26px"></td>
                                <td style="height: 26px"></td>
                            </tr>
                            <tr id="Tr4" runat="server">
                                <td style="width: 281px; height: 26px">Suppress Manager Detail :</td>
                                <td style="height: 26px">
                                    <asp:CheckBox ID="chkSuppressManagerDetail" runat="server" onclick="ClearValues();" /></td>
                                <td style="height: 26px"></td>
                                <td style="height: 26px"></td>
                            </tr>
                            <tr>
                                <td style="height: 36px; width: 281px;">&nbsp;</td>
                                <td style="height: 36px">
                                    <asp:Button ID="btnGetReport" runat="server" OnClick="btnGenerateReport_Click" Text="Generate Report" /></td>
                                <td style="height: 36px">
                                    <asp:DropDownList ID="ddlBatchType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlBatchType_SelectedIndexChanged"
                                        Visible="False">
                                        <asp:ListItem Selected="True" Value="0">Please Select</asp:ListItem>
                                        <asp:ListItem Value="1">MTGBK</asp:ListItem>
                                        <asp:ListItem Value="2">Q</asp:ListItem>
                                        <asp:ListItem Value="3">M</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td style="height: 36px"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="False" TabIndex="1"
                            ToolTip="Batch List" Width="100%" OnRowDataBound="gvList_RowDataBound">
                            <Columns>
                                <asp:BoundField DataField="ssi_hhreportparametersid" HeaderText="ssi_hhreportparametersid"
                                    Visible="False" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" ID="chkbSelectBatch" Checked="false" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField HeaderText="Name" DataField="ssi_name" SortExpression="BatchName" />
                                <asp:BoundField HeaderText="Contact" DataField="Ssi_ContactIdName" SortExpression="Ssi_ContactIdName" />
                                <asp:BoundField HeaderText="Created By" DataField="CreatedByName" SortExpression="CreatedByName" />
                                <asp:BoundField DataField="FolderNameTxt" HeaderText="FolderName" Visible="False" />
                                <asp:BoundField DataField="HouseholdNameTxt" HeaderText="HouseholdNameTxt" Visible="False" />
                                <asp:BoundField DataField="PdfFileName" HeaderText="PdfFileName" Visible="False" />
                                <%-- <asp:BoundField HeaderText="" DataField="" SortExpression="" />
                                <asp:BoundField DataField="" DataFormatString="{0:f3}" HeaderText="" HtmlEncode="False" />
                                <asp:BoundField DataField="" HeaderText="" DataFormatString="{0:f3}" HtmlEncode="False" />--%>
                            </Columns>
                            <HeaderStyle Height="10px" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td align="right">
                        <asp:Button ID="btnGenerateReport" runat="server" OnClientClick="return hidenme(this);"
                            OnClick="btnGenerateReport_Click" Text="Generate Report" Visible="False" />
                        <div id="divdot" style="display: none;">
                            ....
                        </div>

                        <script type="text/javascript">

                            function hidenme(obj) {
                                var isValid = Page_ClientValidate('');
                                if (isValid) {
                                    obj.style.display = "none";
                                    document.getElementById("divdot").style.display = "";
                                    return true;
                                }
                                else {
                                    return false;
                                }
                            }



                        </script>

                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List"
                            ShowMessageBox="True" ShowSummary="False" />
                    </td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
            </table>

            <asp:Chart ID="LineChart1" runat="server" Height="400px" Width="800px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1" Text="Total Investment Assets vs. Net Invested Capital">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="Series1">
                    </asp:Series>
                    <asp:Series Name="Series2">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent" 
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LineWidth="2">
                            <LabelStyle Format="{C0}" />
                            <MajorGrid LineColor="64, 64, 64, 64"  />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LineWidth="2" Interval="Auto">
                            <LabelStyle Format="yyyy" IsEndLabelVisible="true" />
                            <MajorGrid LineColor="64, 64, 64, 64"  Enabled="false"/>
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisX>
                    </asp:ChartArea>
                </ChartAreas>
                <Legends>
                    <asp:Legend LegendStyle="Row" Docking="Bottom" Alignment="Center" TextWrapThreshold="100"
                        AutoFitMinFontSize="7" IsTextAutoFit="false" MaximumAutoSize="100">
                    </asp:Legend>
                </Legends>
            </asp:Chart>
            <asp:Chart ID="LineChart2" runat="server" Height="400px" Width="800px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1" Text="Total Investment Assets vs. Inflation Adj. Net Invested Capital">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="Series1" >
                    </asp:Series>
                    <asp:Series Name="Series2">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent"
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LineWidth="2">
                            <LabelStyle Format="{C0}" />
                            <MajorGrid LineColor="64, 64, 64, 64" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LineWidth="2" Interval="Auto">
                            <LabelStyle Format="yyyy" IsEndLabelVisible="true" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisX>
                    </asp:ChartArea>
                </ChartAreas>
                <Legends>
                    <asp:Legend LegendStyle="Row" Docking="Bottom" Alignment="Center" TextWrapThreshold="100"
                        AutoFitMinFontSize="7" IsTextAutoFit="false" MaximumAutoSize="100">
                    </asp:Legend>
                </Legends>
            </asp:Chart>
            <asp:Chart ID="LineChart" runat="server" Height="195px" Width="1200px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1" Text="Growth of My Gresham Advised Assets (GAA)">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="Series1">
                    </asp:Series>
                    <asp:Series Name="Series2">
                    </asp:Series>
                    <asp:Series Name="Series3">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent"
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LineWidth="2">
                            <LabelStyle Format="{C0}" />
                            <MajorGrid LineColor="#868686" LineWidth="2"  LineDashStyle="Solid"   />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LineWidth="2" Interval="Auto">
                            <LabelStyle Format="yyyy" IsEndLabelVisible="true" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisX>
                    </asp:ChartArea>
                </ChartAreas>
                <Legends>
                    <asp:Legend LegendStyle="Row" Docking="Bottom" Alignment="Center" TextWrapThreshold="100"
                        AutoFitMinFontSize="7" IsTextAutoFit="false" MaximumAutoSize="100">
                    </asp:Legend>
                </Legends>
            </asp:Chart>
            <asp:Chart ID="Chart1" runat="server" Height="195px" Width="1200px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1" Text="Annual Performance of Gresham Advised Assets (GAA)">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="RETURN" ChartType="Column" IsXValueIndexed="false" Color="#2A6FB6"
                        BorderColor="Transparent">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent"
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LabelAutoFitMaxFontSize="8" LineWidth="2">
                            <LabelStyle Font="Frutiger55, 2pt, style=regular" Format="{0.0}%" IsEndLabelVisible="true" />
                            <MajorGrid LineColor="64, 64, 64, 64" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LabelAutoFitMaxFontSize="8" LineWidth="2" Interval="1">
                            <LabelStyle Format="yyyy" />
                            <MajorGrid LineColor="64, 64, 64, 64" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" Interval="1" LineDashStyle="Solid" />
                        </AxisX>
                    </asp:ChartArea>
                </ChartAreas>
            </asp:Chart>

            <asp:Chart ID="ShapeChartRpt4" Height="750px" Width="750px" runat="server" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1" Text="Performance vs. Volatility (since 01/01/2011)">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="Series1">
                    </asp:Series>
                    <asp:Series Name="Series2">
                    </asp:Series>
                    <asp:Series Name="Series3">
                    </asp:Series>
                    <asp:Series Name="Series4">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent"
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LineWidth="2" Title="Annualized Return %" TitleFont="Frutiger55, 9pt, style=Bold">
                            <LabelStyle Format="{N0}%" Font="Frutiger55, 8pt" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LineWidth="2" TitleAlignment="Center" Interval="Auto" TitleFont="Frutiger55, 9pt, style=Bold" Title="Annualized Volatility (Standard Deviation)">
                            <LabelStyle Format="{N0}%" IsEndLabelVisible="true" Font="Frutiger55, 8pt" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisX>
                    </asp:ChartArea>
                </ChartAreas>
            </asp:Chart>
            <asp:Chart ID="ColumnChartRpt4" Height="750px" Width="750px" runat="server" Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Alignment="TopCenter" Name="Title1" Text="Portfolio Protection During Worst Market Months">
                    </asp:Title>
                </Titles>
                <Series>
                    <asp:Series Name="Series1" Color="#558ED5"></asp:Series>
                </Series>
                <Series>
                    <asp:Series Name="Series2" Color="#ffffff"></asp:Series>
                </Series>
                <Series>
                    <asp:Series Name="Series3" Color="#003399"></asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent"
                        BackColor="Transparent" ShadowColor="Transparent">
                        <AxisY LineColor="#868686" LineWidth="2" TitleFont="Frutiger55, 9pt, style=Bold">
                            <LabelStyle Format="{N0}%" Font="Frutiger55, 6pt" IsEndLabelVisible="true" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisY>
                        <AxisX LineColor="#868686" LineWidth="2" Interval="Auto" TitleFont="Frutiger55, 9pt, style=Bold" Title="Return %">
                            <LabelStyle Font="Frutiger55, 8pt, style=Bold" />
                            <MajorGrid LineColor="64, 64, 64, 64" Enabled="false" />
                            <MinorTickMark LineColor="64, 64, 64, 64" LineWidth="2" Size="1" LineDashStyle="Solid" />
                        </AxisX>

                    </asp:ChartArea>
                </ChartAreas>
                <Legends>
                    <asp:Legend LegendStyle="Row" Docking="Bottom" TextWrapThreshold="100" TitleFont="Frutiger55, 8pt"
                        AutoFitMinFontSize="7" IsTextAutoFit="false" MaximumAutoSize="100">
                    </asp:Legend>
                </Legends>
            </asp:Chart>
        </div>
    </form>

    <script language="javascript" type="text/javascript">
        function ClearLabel() {
            document.getElementById('<%= lblError.ClientID%>').innerHTML = "";
    }

    </script>

</body>
</html>
