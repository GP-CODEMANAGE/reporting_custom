<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PerfAnalyticsTEST.aspx.cs" Inherits="PerfAnalytics1"
    Culture="en-US" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Gresham Advised Assets Performance Summary Report GA 2.0</title>
    <link id="style1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        function ValidateDate() {
            var Household = document.getElementById("ddlHousehold").value;
            var Group = document.getElementById("ddlGroup").value;
            var AsofDate = document.getElementById("txtAsofdate").value;
            //var AllocationGrp = document.getElementById("ddlAllocationGrp").value;
            var TIAGrp = document.getElementById("ddlTIAGrp").value;
            var chkrpt1 = document.getElementById('chkrpt1');
            var chkrpt3 = document.getElementById('chkrpt3');
            var chkrpt4 = document.getElementById('chkrpt4');

            if (Household == "") {
                alert("Please Select HouseHold.");
                return false;
            }

            //            if (AllocationGrp == "") {
            //                alert("Please Select Allocation Group.");
            //                return false;
            //            }

            if (TIAGrp == "" && Group == "") {
                alert("Please Select TIA Group Or Group Name");
                return false;
            }

            //            if (Group == "") {
            //                alert("Please Select Group.");
            //                return false;
            //            }
            if (AsofDate == "") {
                alert("Please enter As Of Date.");
                return false;
            }

            if (chkrpt1.checked == false && chkrpt3.checked == false && chkrpt4.checked == false) {
                alert("please select atleast one report");
                return false;
            }

        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
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
                                <td colspan="3" align="left">
                                    <asp:Label ID="lblHeader" runat="server" Font-Bold="True" Font-Size="Large" Text="Perf Analytics"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 18px" valign="top" colspan="3">
                                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%; height: 26px;">
                                    <asp:Label ID="lblHouseHold" runat="server" Text="HouseHold"></asp:Label>
                                </td>
                                <td style="width: 80%; height: 26px;">
                                    <asp:DropDownList ID="ddlHousehold" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlHousehold_SelectedIndexChanged">
                                    </asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="ddlHousehold"
                                        Display="None" ErrorMessage="Please Select HouseHold"></asp:RequiredFieldValidator>
                                </td>
                                <td style="width: 4px; height: 26px;"></td>
                            </tr>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label11" runat="server" Text="Group:"></asp:Label>
                                </td>
                                <td style="height: 40px">
                                    <asp:DropDownList ID="ddlGroup" runat="server">
                                    </asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="ddlGroup"
                                        Display="None" ErrorMessage="Please Select Group"></asp:RequiredFieldValidator>
                                </td>
                                <td style="width: 4px; height: 40px"></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label3" runat="server" Text="As Of Date:"></asp:Label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtAsofdate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtAsofdate)">
                                        <img id="imgorgDateRec" alt="" border="0" src="images/calander.png" /></a>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsofdate"
                                        Display="None" ErrorMessage="Please enter As Of Date"></asp:RequiredFieldValidator><asp:CustomValidator
                                            ID="CustomValidator1" runat="server" ControlToValidate="txtAsofdate" ErrorMessage="As of date is not valid"
                                            ClientValidationFunction="ValidateForm" Display="None"> </asp:CustomValidator>
                                </td>
                                <td style="width: 4px"></td>
                            </tr>
                            <%--<tr>
                            <td style="width: 25%">
                                <asp:Label ID="Label2" runat="server" Text="AllocationGroup:"></asp:Label>
                            </td>
                            <td style="height: 40px">
                                <asp:DropDownList ID="ddlAllocationGrp" runat="server">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="ddlAllocationGrp"
                                    Display="None" ErrorMessage="Please Select Allocation Group"></asp:RequiredFieldValidator>
                            </td>
                            <td style="width: 4px; height: 40px">
                            </td>
                        </tr>--%>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label4" runat="server" Text="TIA Group:"></asp:Label>
                                </td>
                                <td style="height: 40px">
                                    <asp:DropDownList ID="ddlTIAGrp" runat="server">
                                    </asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="ddlTIAGrp"
                                        Display="None" ErrorMessage="Please Select TIA Group"></asp:RequiredFieldValidator>
                                </td>
                                <td style="width: 4px; height: 40px"></td>
                            </tr>
                            <%--   <tr>
                                <td style="width: 20%">
                                    <asp:Label ID="Label2" runat="server" Text="Start Date:"></asp:Label></td>
                                <td>
                                    <asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>&nbsp;&nbsp;<a onclick="showCalendarControl( txtStartDate)">
                                        <img id="img1" alt="" border="0" src="images/calander.png" /></a>
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>--%>
                            <tr>
                                <td style="width: 25%">
                                    <asp:Label ID="Label1" runat="server" Text="Asset Class:"></asp:Label>
                                </td>
                                <td style="height: 40px">
                                    <asp:ListBox ID="lstAssetClass" runat="server" Height="170px" SelectionMode="Multiple"></asp:ListBox>
                                </td>
                                <td style="width: 4px; height: 40px;"></td>
                            </tr>
                            <tr>
                                <td colspan="3">&nbsp;</td>
                            </tr>
                            <tr>
                                <td></td>
                                <td colspan="2">

                                    <asp:CheckBox ID="chkrpt1" Checked="true" Text=" TIA Charts" runat="server" AutoPostBack="true" /></td>


                            </tr>
                            <tr>
                                <td></td>
                                <td colspan="2">
                                    <asp:CheckBox ID="chkrpt3" Checked="true" Text="Absolute Returns" runat="server" AutoPostBack="true" />
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                                <td colspan="2">
                                    <asp:CheckBox ID="chkrpt4" Checked="true" Text="Capital Protection" runat="server" AutoPostBack="true" />
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                                <td valign="top">
                                    <br />
                                    <asp:Button ID="Button1" runat="server" Text="Generate Report" OnClientClick="return ValidateDate();"
                                        OnClick="Button1_Click" />
                                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
                                        ShowSummary="False" />
                                </td>
                                <td style="width: 4px"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
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
            <%--<asp:Chart ID="LineChart" runat="server" Height="195px" Width="1200px" BorderlineDashStyle="Solid"
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
            </asp:Chart>--%>
           <%-- <asp:Chart ID="LineChart1" runat="server" Height="400px" Width="800px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1"  Text="Total Investment Assets vs. Net Invested Capital">
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
            </asp:Chart>--%>
            <asp:Chart ID="LineChart2" runat="server" Height="400px" Width="800px" BorderlineDashStyle="Solid"
                Visible="false">
                <Titles>
                    <asp:Title Visible="false" Font="Frutiger55, 9pt, style=Bold" Name="Title1"   Text="Total Investment Assets vs. Inflation Adj. Net Invested Capital">
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
                        <AxisX LineColor="#868686" LineWidth="2" TitleAlignment="Center"  Interval="Auto" TitleFont="Frutiger55, 9pt, style=Bold" Title="Annualized Volatility (Standard Deviation)">
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
                            <LabelStyle Format="{N0}%" Font="Frutiger55, 6pt" />
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
</body>
</html>
