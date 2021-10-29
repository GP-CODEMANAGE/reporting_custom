<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ClientServicesDashboard.aspx.cs" Inherits="ClientServicesDashboard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <%--<META HTTP-EQUIV='Pragma' CONTENT='no-cache'">
<META HTTP-EQUIV="Expires" CONTENT="-1">--%>
    <title>ClientServicesDashboard</title>
    <%--<link href="bootstrap/css/bootstrap.css" rel="stylesheet" />
    <script src="bootstrap/js/bootstrap.js" type="text/javascript"></script>--%>
    <link href="common/gresham.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js" type="text/javascript"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" type="text/javascript"></script>
    <style type="text/css">
        .style1 
        {
            width: 12px;
        }

        .auto-style34
        {
            width: 1000px;
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
            font-size: 18pt;
            font-weight: normal;
            text-decoration: none;
        }

        #myBtn
        {
            display: none;
            position: fixed;
            bottom: 20px;
            right: 30px;
            z-index: 99;
            font-size: 18px;
            border: none;
            outline: none;
            background-color: #006699;
            color: white;
            cursor: pointer;
            padding: 15px;
            border-radius: 4px;
        }

            #myBtn:hover
            {
                background-color: #555;
            }

        .auto-style35
        {
            height: 81px;
        }

        .gvstyling th
        {
            /*background-color: Red;*/
            font-size: 10px;
            text-align: center;
        }

        .gvStyling body
        {
            width: 1300px;
        }



        .DropDown
        {
            font-family: Verdana;
            font-size: x-small;
        }

        .GridViewClass td, th
        {
            padding-right: 10px;
            padding-left: 10px;
        }

        .GridViewClass1 td, th
        {
            /*//padding-right: 10px;*/
            padding-left: 10px;
            padding-right: 10px;
        }

        .grideNewWidth
        {
            width: 100%;
            /*width: 100%;*/
            /*height: 100%;*/
        }

        .PanelNew
        {
            width: 1900px;
            /* height:300px;*/
        }

        .divHeightScroll
        {
            height: 400px;
            overflow: scroll;
        }

        .FixedHeader
        {
            /*position: absolute;*/
        }
    </style>

    <script type="text/javascript">
        // When the user scrolls down 20px from the top of the document, show the button
        window.onscroll = function () { scrollFunction() };

        function scrollFunction() {
            if (document.body.scrollTop > 20 || document.documentElement.scrollTop > 20) {
                document.getElementById("myBtn").style.display = "block";
            } else {
                document.getElementById("myBtn").style.display = "none";
            }
        }

        // When the user clicks on the button, scroll to the top of the document
        function topFunction() {
            document.body.scrollTop = 0;
            document.documentElement.scrollTop = 0;
        }
    </script>



</head>
<body>
    <button onclick="topFunction()" id="myBtn" title="Go to top">Top</button>
    <form id="form1" runat="server" style="width: auto">
        <div onload="LOAD()">


            <table class="auto-style34">
                <%--   <tr>
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
            </tr>--%>
                <tr>
                    <td colspan="3">
                        <img src="images/Gresham_Logo__.jpg" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3">Gresham Partners, LLC
                    </td>
                </tr>
                <tr>
                    <td></td>

                </tr>
                <tr>
                    <td></td>
                </tr>

                <tr>
                    <td>
                        <div id="Div1">
                            <asp:Label ID="lblerror" runat="server"></asp:Label>
                        </div>
                    </td>
                </tr>

                <tr>
                    <td>
                        <div id="id1" style="text-align: center">
                            <asp:Label ID="lblHoushold" runat="server" CssClass="Titlebig"></asp:Label>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <br />

                    </td>
                </tr>

                <tr>
                    <%--<td style="border-bottom: gray 1px solid; text-align: center; height:30px" colspan="2">&nbsp;
                    </td>
                    <td style="width: 5px; height:30px;"></td>--%>
                </tr>
                <tr>
                    <td class="auto-style35">
                        <%--<div class="center">--%>
                        <div id="">
                            <asp:Panel ID="Panel1" runat="server" CssClass="panel panel-default PanelNew" BackColor="#006699">
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                              <asp:HyperLink ID="HyperLink4" href="#LE" runat="server" ForeColor="White">Legal Entity</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:HyperLink ID="HyperLink7" href="#MoneyMovement" runat="server" ForeColor="White">Money Movement</asp:HyperLink>

                                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                            <asp:HyperLink ID="HyperLink1" href="#Recommendation" runat="server" ForeColor="White">Recommendation</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                           <asp:HyperLink ID="HyperLink6" href="#Sales&Purchase" runat="server" ForeColor="White">Purchase & Sale</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                           <asp:HyperLink ID="HyperLink2" href="#CallReports" runat="server" ForeColor="White">Activities</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                           <asp:HyperLink ID="HyperLink3" href="#Task" runat="server" ForeColor="White">Task</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                          <asp:HyperLink ID="HyperLink8" href="#Email" runat="server" ForeColor="White">Email</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                         <asp:HyperLink ID="HyperLink5" href="#Allocation" runat="server" ForeColor="White">Target Vs. Actual Allocation</asp:HyperLink>

                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                       

                        

                            </asp:Panel>
                        </div>
                    </td>

                </tr>
                <tr>
                    <%-- <td style="border-top: gray 1px solid; text-align: center; height: 30px;" colspan="2">&nbsp;
                    </td>
                    <td style="width: 2px; height: 30px;"></td>--%>
                </tr>

                <tr>
                    <td>
                        <br />
                    </td>
                </tr>
            </table>


            <table>
                <asp:ScriptManager ID="ScriptManager1" runat="server">
                </asp:ScriptManager>
                <tr>
                    <%--LegalEntity panel--%>
                    <td>
                        <div id="LE">
                            <asp:Panel ID="PanelLE" runat="server" CssClass="panel panel-default PanelNew">

                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4 class="LE">Legal Entity  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblLETimeStamp" runat="server" Text=""></asp:Label>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                  
                                    
                                    </h4>
                                </div>
                                <%-- <div id="collapse2" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <%--<asp:LinkButton ID="lbLE" runat="server" Text="Export To Excel" OnClick="lbLE_Click">--%><%--<img src="images/Picture1.png"/>Export to Excel--%></asp:LinkButton>
                                    <asp:Label ID="lblAlertLE" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanelLE" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvLE" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            OnSorting="gvLE_Sorting" AllowSorting="true" HorizontalAlign="center" CssClass="gvstyling th grideNewWidth GridViewClass td, th" CellPadding="2" BorderWidth="2px" BorderColor="#CCCCCC">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" /> ItemStyle-HorizontalAlign="Right" --%>
                                                                <asp:BoundField DataField="ssi_name" HeaderText="Name" SortExpression="ssi_name" />
                                                                <%--0--%>

                                                                <asp:BoundField DataField="investortype" HeaderText="Investor Type" SortExpression="investortype" />
                                                                <%--1--%>

                                                                <asp:BoundField DataField="CapitalCallAccount" HeaderText="Capital Call Account" SortExpression="CapitalCallAccount" />
                                                                <%--2--%>

                                                                <asp:BoundField DataField="DistributionAccount" HeaderText="Distribution Account" SortExpression="DistributionAccount"/>
                                                                <%--3--%>

                                                                <asp:BoundField DataField="Signor1" HeaderText="Signer" SortExpression="Signor1" />
                                                                <%--4--%>

                                                                <asp:BoundField DataField="CreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--5--%>

                                                               <%-- <asp:BoundField DataField="Signor3" HeaderText="Signer 3" SortExpression="Signor3" />--%>
                                                                <%--6--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" />
                                                        </asp:GridView>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <%-- </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>
                <tr>
                    <%--Money Market Moment--%>
                    <td>
                        <div id="MoneyMovement">
                            <asp:Panel ID="Panel3" runat="server" CssClass="panel panel-default PanelNew">

                                <%--<div class="center" id="Recommendation">--%>
                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4>Money Movement 
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblMMTimeStamp" runat="server" Text=""></asp:Label>

                                    </h4>
                                </div>
                                <%--  <div id="collapse1" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <div id="Div4" style="text-align: right">
                                        <%--  <asp:LinkButton ID="lbExporttoExcel" runat="server" Text="Export To Excel" OnClick="lbExporttoExcel_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%></asp:LinkButton>
                                    </div>
                                    <asp:Label ID="lblMM" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanel3" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvMoney" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            AllowSorting="true" CssClass="gvstyling th GridViewClass td, th grideNewWidth" BorderWidth="2px" BorderColor="#CCCCCC" OnSorting="gvMoney_Sorting">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                                                                <asp:BoundField DataField="LegalEntity" HeaderText="Legal Entity" SortExpression="LegalEntity" />
                                                                <%--0--%>

                                                                <asp:BoundField DataField="Custodian" HeaderText="Custodian" SortExpression="Custodian" />
                                                                <%--1--%>

                                                                <asp:BoundField DataField="Account" HeaderText="Account" SortExpression="Account" />
                                                                <%--2--%>

                                                                <asp:BoundField DataField="DispTradeDate" HeaderText="Trade Date" SortExpression="DispTradeDate" />
                                                                <%--3--%>

                                                                <asp:BoundField DataField="security" HeaderText="Security" SortExpression="security" />
                                                                <%--3--%>

                                                                <asp:BoundField DataField="TradeAmount" HeaderText="Trade Amount" SortExpression="TradeAmount" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:C0}" />
                                                                <%--4--%>

                                                                <asp:BoundField DataField="GreshamAdvised" HeaderText="GA" SortExpression="GreshamAdvised" />
                                                                <%--5--%>

                                                              <%--  <asp:BoundField DataField="TransactionCode" HeaderText="Transaction Type" SortExpression="TransactionCode" />--%>
                                                                <%--6--%>
                                                                <asp:BoundField DataField="Fund" HeaderText="Fund" SortExpression="Fund" />
                                                                <%--7--%>
                                                                <asp:BoundField DataField="Comment" HeaderText="Comment" SortExpression="Comment" ItemStyle-Width="500px" />
                                                                <%--8--%>

                                                                <asp:BoundField DataField="DispCreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--9--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" CssClass="FixedHeader" />
                                                        </asp:GridView>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>

                                    <%--  </div>--%>




                                    <%--  </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>
                <tr>
                    <%--Recommendation panel--%>
                    <td>
                        <div id="Recommendation">
                            <asp:Panel ID="PanelRecommendation" runat="server" CssClass="panel panel-default PanelNew">

                                <%--<div class="center" id="Recommendation">--%>
                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4>Recommendation &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblRecommTimeStamp" runat="server" Text=""></asp:Label></h4>
                                </div>
                                <%--  <div id="collapse1" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <div id="" style="text-align: right">
                                        <%--  <asp:LinkButton ID="lbExporttoExcel" runat="server" Text="Export To Excel" OnClick="lbExporttoExcel_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%></asp:LinkButton>
                                    </div>
                                    <asp:Label ID="lblalertRecommendation" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvRecommendation" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            OnSorting="gvRecommendation_Sorting" AllowSorting="true" CssClass="gvstyling th GridViewClass td, th grideNewWidth" BorderWidth="2px" BorderColor="#CCCCCC">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>

                                                                <asp:BoundField DataField="CreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--0--%>
                                                                <asp:BoundField DataField="ssi_closedateinvestment" HeaderText="Close Date of Investment" SortExpression="ssi_closedateinvestment" />
                                                                <%--1--%>

                                                                <asp:BoundField DataField="ssi_legalentityidname" HeaderText="Legal Entity" SortExpression="ssi_legalentityidname" />
                                                                <%--2--%>

                                                                <asp:BoundField DataField="ssi_fundidname" HeaderText="Fund" SortExpression="ssi_fundidname" />
                                                                <%--3--%>

                                                                <asp:BoundField DataField="ssi_proposedamount" HeaderText="Proposed Amount" SortExpression="ssi_proposedamount" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:C0}" />
                                                                <%--4--%>

                                                                <asp:BoundField DataField="ssi_confirmedamount" HeaderText="Confirmed Amount" SortExpression="ssi_confirmedamount" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:C0}" />
                                                                <%--5--%>

                                                                <asp:BoundField DataField="transactiontypes" HeaderText="Transaction Types" SortExpression="transactiontypes" />
                                                                <%--6--%>

                                                                <asp:BoundField DataField="status" HeaderText="Status" SortExpression="status" />
                                                                <%--7 --%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" CssClass="FixedHeader" />
                                                        </asp:GridView>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>

                                    <%--  </div>--%>




                                    <%--  </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>
                <tr>
                    <%--Sales&Purchase panel--%>
                    <td>
                        <div id="Sales&Purchase">
                            <asp:Panel ID="Panel2" runat="server" CssClass="panel panel-default PanelNew">

                                <%--<div class="center" id="Recommendation">--%>
                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4>Purchase & Sale &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <asp:Label ID="lblSales" runat="server" Text=""></asp:Label>

                                    </h4>
                                </div>
                                <%--  <div id="collapse1" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <div id="Div3" style="text-align: right">
                                        <%--  <asp:LinkButton ID="lbExporttoExcel" runat="server" Text="Export To Excel" OnClick="lbExporttoExcel_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%></asp:LinkButton>
                                    </div>
                                    <asp:Label ID="lblSP" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvSalesPurchase" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            AllowSorting="true" CssClass="gvstyling th GridViewClass td, th grideNewWidth" BorderWidth="2px" BorderColor="#CCCCCC" OnSorting="gvSalesPurchase_Sorting">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>

                                                                <asp:BoundField DataField="LegalEntity" HeaderText="Legal Entity" SortExpression="LegalEntity" />
                                                                <%--0--%>

                                                                <asp:BoundField DataField="DispTradeDate" HeaderText="Trade Date" SortExpression="DispTradeDate" />
                                                                <%--1--%>

                                                                <asp:BoundField DataField="TransactionCode" HeaderText="Transaction Code Description" SortExpression="TransactionCode" />
                                                                <%--2--%>

                                                                <asp:BoundField DataField="Quantity" HeaderText="Quantity" SortExpression="Quantity" DataFormatString="{0:N0}" />
                                                                <%--3--%>

                                                                <asp:BoundField DataField="security" HeaderText="Security" SortExpression="security" />
                                                                <%--4--%>


                                                                <asp:BoundField DataField="Fund" HeaderText="Fund" SortExpression="Fund" />
                                                                <%--5--%>


                                                                <asp:BoundField DataField="TradeAmount" HeaderText="Trade Amount" SortExpression="TradeAmount" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:C0}" />
                                                                <%--6--%>


                                                                <asp:BoundField DataField="GreshamAdvised" HeaderText="GA" SortExpression="GreshamAdvised" />
                                                                <%--7--%>

                                                                <asp:BoundField DataField="DispCreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--8--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" CssClass="FixedHeader" />
                                                        </asp:GridView>
                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>

                                    <%--  </div>--%>




                                    <%--  </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>
                <tr>
                    <%--Activity CallReport panel--%>
                    <td>
                        <div id="CallReports">
                            <asp:Panel ID="PanelCallReports" runat="server" CssClass="panel panel-default PanelNew">

                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4 class="CallReports">Activities &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblActivityTimeFrame" runat="server" Text=""></asp:Label>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                 
                                    </h4>
                                </div>
                                <%-- <div id="collapseCallReport" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <%--  <asp:LinkButton ID="LinkbtnCallReport" runat="server" Text="Export To Excel" OnClick="LinkbtnCallReport_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%><%--</asp:LinkButton>--%>
                                    <asp:Label ID="lblAlertCallReport" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanelCallReports" runat="server">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvcallreports" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            OnSorting="gvcallreports_Sorting" AllowSorting="true" CssClass="gvstyling th grideNewWidth" CellPadding="2" BorderWidth="2px" BorderColor="#CCCCCC">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                                                                <%-- <asp:BoundField DataField="Subject" HeaderText="Subject" SortExpression="Subject" />--%>
                                                                <asp:HyperLinkField DataNavigateUrlFields="LinkUrl" DataTextField="Subject"
                                                                    HeaderText="Subject" ControlStyle-CssClass="localLink" SortExpression="Subject" Target="_blank" />
                                                                <%--0--%>
                                                              <%--  <asp:BoundField DataField="ssi_from" HeaderText="Call From" SortExpression="ssi_from" />--%>
                                                                <%--1--%>
                                                                <asp:BoundField DataField="ssi_to" HeaderText="Participants/Regarding" SortExpression="ssi_to" />
                                                                <%--2--%>
                                                                <%--  <asp:BoundField DataField="ssi_notes" HeaderText="Notes" SortExpression="ssi_notes" />--%>
                                                                <%--3--%>
                                                                <asp:BoundField DataField="directioncode" HeaderText="Direction" SortExpression="directioncode" />
                                                                <%--4--%>
                                                                <asp:BoundField DataField="scheduledend" HeaderText="Scheduled Follow Up" SortExpression="scheduledend" />
                                                                <%--5--%>
                                                                <asp:BoundField DataField="statecode" HeaderText="Activity Status" SortExpression="statecode" />

                                                                <%--6--%>
                                                                <asp:BoundField DataField="OwnerIdName" HeaderText="Owner" SortExpression="OwnerIdName" />
                                                                <%--7--%>
                                                                <asp:BoundField DataField="Type" HeaderText="Type" SortExpression="Type" />
                                                                <%--8--%>
                                                                <asp:BoundField DataField="CreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--9--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" CssClass="FixedHeader" />
                                                        </asp:GridView>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <%-- </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>
                <tr>
                    <%--Task panel--%>
                    <td>
                        <div id="Task">
                            <asp:Panel ID="PanelTask" runat="server" CssClass="panel panel-default PanelNew">


                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4 class="Task">Task &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblTaskTimeFrame" runat="server" Text=""></asp:Label>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                 
                                    </h4>
                                </div>
                                <%-- <div id="collapseTask" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <%-- <asp:LinkButton ID="Linkbtntask" runat="server" Text="Export To Excel" OnClick="Linkbtntask_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%></asp:LinkButton>
                                    <asp:Label ID="lblAlertTask" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanelTask" runat="server">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvTask" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            Style="margin-top: 0px" OnSorting="gvTask_Sorting" AllowSorting="true" CssClass="gvstyling th grideNewWidth" CellPadding="2" BorderWidth="2px" BorderColor="#CCCCCC">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                                                                <%-- <asp:BoundField DataField="Subject" HeaderText="Subject" SortExpression="Subject" ItemStyle-Width="800px" />--%>
                                                                <asp:HyperLinkField DataNavigateUrlFields="LinkUrl" DataTextField="Subject"
                                                                    HeaderText="Subject" ControlStyle-CssClass="localLink" SortExpression="Subject" ItemStyle-Width="800px" Target="_blank" />
                                                                <%--0--%>
                                                                <%-- <asp:BoundField DataField="ssi_notes" HeaderText="Notes" SortExpression="ssi_notes" />--%>
                                                                <%--1--%>
                                                                <asp:BoundField DataField="PriorityCode" HeaderText="Priority Code" SortExpression="PriorityCode" />
                                                                <%--2--%>
                                                                <asp:BoundField DataField="DueDate" HeaderText="Due Date" SortExpression="DueDate" />
                                                                <%--3--%>
                                                                <asp:BoundField DataField="statecode" HeaderText="Status" SortExpression="statecode" />
                                                                <%--4--%>
                                                                <asp:BoundField DataField="OwnerIdName" HeaderText="Owner" SortExpression="OwnerIdName" />
                                                                <%--5--%>
                                                                <asp:BoundField DataField="CreatedOn" HeaderText="Created On" SortExpression="CreatedOn" />
                                                                <%--6--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" />
                                                        </asp:GridView>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <%-- </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>




                 <tr>
                    <%--Email panel--%>
                    <td>
                        <div id="Email">
                            <asp:Panel ID="PanelEmail" runat="server" CssClass="panel panel-default PanelNew">


                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4 class="Email">Email &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="lblEmailTimeFrame" runat="server" Text=""></asp:Label>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                 
                                    </h4>
                                </div>
                                <%-- <div id="collapseTask" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <%-- <asp:LinkButton ID="LinkbtnEmail" runat="server" Text="Export To Excel" OnClick="Linkbtntask_Click">--%><%--<img src="images/Picture1.png" />Export to Excel--%></asp:LinkButton>
                                    <asp:Label ID="lblAlertEmail" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive divHeightScroll">
                                                <asp:UpdatePanel ID="UpdatePanelEmail" runat="server">
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvEmail" runat="server"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            Style="margin-top: 0px" OnSorting="gvEmail_Sorting" AllowSorting="true" CssClass="gvstyling th grideNewWidth" CellPadding="2" BorderWidth="2px" BorderColor="#CCCCCC">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                                                                <%-- <asp:BoundField DataField="Subject" HeaderText="Subject" SortExpression="Subject" ItemStyle-Width="800px" />--%>
                                                                <asp:HyperLinkField DataNavigateUrlFields="LinkUrl" DataTextField="Subject"
                                                                    HeaderText="Subject" ControlStyle-CssClass="localLink" SortExpression="Subject" ItemStyle-Width="600px" Target="_blank" />
                                                                <%--0--%>                                                               
                                                                <asp:BoundField DataField="Sender" HeaderText="Sender" SortExpression="Sender" />
                                                                <%--1--%>
                                                                <asp:BoundField DataField="Recipients" HeaderText="Recipients" SortExpression="Recipients" />
                                                                <%--2--%>
                                                                <asp:BoundField DataField="StatusReason" HeaderText="StatusReason" SortExpression="StatusReason" />
                                                                <%--3--%>
                                                                <%--<asp:BoundField DataField="TimeFrame" HeaderText="TimeFrame" SortExpression="TimeFrame" />--%>
                                                                <%--4--%>
                                                                <asp:BoundField DataField="ModifiedOn" HeaderText="Modified On" SortExpression="ModifiedOn" />
                                                                <%--5--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" />
                                                        </asp:GridView>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <%-- </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>












                <tr>
                    <%--Allocation panel--%>
                    <td>
                        <div id="Allocation">
                            <asp:Panel ID="PanelAllocation" runat="server" CssClass="panel panel-default PanelNew">
                                <div class="panel-heading" style="padding-top: 0px; padding-bottom: 1px">
                                    <h4 class="Allocation">Target Vs. Actual Allocation
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

                                        <asp:DropDownList ID="ddlList" runat="server" CssClass="DropDown" OnSelectedIndexChanged="ddlList_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <asp:Label ID="lblAsofdate" runat="server"></asp:Label>
                                         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <%--<asp:Label ID="lblTagetTimeFrame" runat="server" Text=""></asp:Label>--%>
                                    </h4>
                                </div>
                                <%--<div id="collapseAllocation" class="panel-collapse collapse">--%>
                                <div class="panel-body">
                                    <%-- <asp:LinkButton ID="LinkbtnAllocation" runat="server" Text="Export To Excel" OnClick="LinkbtnAllocation_Click">--%><%--<img src="images/Picture1.png"/>Export to Excel--%></asp:LinkButton>
                                    <asp:Label ID="lblAlertAllocation" runat="server" ForeColor="Red"></asp:Label>
                                    <div class="row">
                                        <div class="col-lg-12 ">
                                            <div class="table-responsive">
                                                <asp:UpdatePanel ID="UpdatePanelAllocation" runat="server" UpdateMode="Conditional">
                                                    <Triggers>
                                                        <asp:AsyncPostBackTrigger ControlID="ddlList" EventName="SelectedIndexChanged" />
                                                    </Triggers>
                                                    <ContentTemplate>
                                                        <asp:GridView ID="gvAllocation" runat="server" Width="1300px"
                                                            AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="Smaller"
                                                            Style="margin-top: 0px" OnSorting="gvAllocation_Sorting" AllowSorting="true" GridLines="None" CssClass="gvstyling th GridViewClass1 td, th grideNewWidth" CellPadding="2" BorderWidth="1" BorderColor="#006699">
                                                            <Columns>
                                                                <%--  <asp:BoundField DataField="AccountLegalEntityName" HeaderText="Legal Entity Name" />--%>
                                                                <asp:BoundField DataField="Asset Class" HeaderText="Asset Class" SortExpression="Asset Class" ItemStyle-Width="300px" />
                                                                <%--0--%>
                                                                <asp:BoundField DataField="Current Portfolio Value" HeaderText="Current Allocation" SortExpression="Current Portfolio Value" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:C0}" ItemStyle-Width="140px" />
                                                                <%--1--%>
                                                                <asp:BoundField DataField="Current Portfolio %" HeaderText="Current Portfolio %" SortExpression="Current Portfolio %" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:P1}" ItemStyle-Width="140px" />
                                                                <%--2--%>
                                                                <asp:BoundField DataField="Suggested Allocation" HeaderText="Suggested Allocation" SortExpression="Suggested Allocation" ItemStyle-HorizontalAlign="Right" DataFormatString="{0:P1}" ItemStyle-Width="140px" />
                                                                <%--3--%>
                                                                <asp:BoundField DataField="Bench mark" HeaderText="Bench mark" SortExpression="Bench mark" ItemStyle-Width="400px" />
                                                                <%--4--%>
                                                            </Columns>
                                                            <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                                            <%--  <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />--%>
                                                            <RowStyle ForeColor="#000066" />
                                                            <HeaderStyle Height="10px" />
                                                        </asp:GridView>

                                                    </ContentTemplate>
                                                </asp:UpdatePanel>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <%--  </div>--%>
                            </asp:Panel>
                        </div>
                    </td>
                </tr>



              






            </table>

        </div>

    </form>
</body>
</html>

