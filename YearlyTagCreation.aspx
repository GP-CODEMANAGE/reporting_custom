<%@ Page Language="C#" AutoEventWireup="true" CodeFile="YearlyTagCreation.aspx.cs" Inherits="YearlyTagCreation" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
      <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    

    <title></title>
    <style type="text/css">
        .Titlebig {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: normal;
            text-decoration: none;
        }

        span {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }

        input, select {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }


        .style3 {
            width: 11%;
        }


        .style4 {
            width: 1%;
        }


        </style>


    
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width: 100%">
            <tr>
                <td colspan="3">
                    <img src="images/Gresham_Logo__.jpg" />
                </td>
            </tr>
            <tr>
                <td colspan="3" class="Titlebig">
                    Gresham Partners, LLC
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <b>Yearly Tag Creation</b></td>
            </tr>
            <tr>
                <td style="height: 18px" valign="top" colspan="3">
                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                            <td style="width: 25%; height: 24px;">
                                Please Select the Year to Create Tag</td>
                            <td style="height: 24px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlYear" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlYear_SelectedIndexChanged"
                                    >
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 24px;">
                            </td>
                        </tr>
         
           <tr>
                            <td style="width: 25%; height: 24px;">
                                Please Select the Folder </td>
                            <td style="height: 24px">
                                <asp:DropDownList Font-Names="Verdana" ID="ddlFolderName" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlFolderName_SelectedIndexChanged"
                                    >
                                </asp:DropDownList></td>
                            <td style="width: 4px; height: 24px;">
                            </td>
                        </tr>
         
          
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click" />
                    <br />
                    <br />
                    <asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label>
                    <br />
                    <asp:Label ID="lblMsg0" runat="server" ForeColor="Red"></asp:Label>
                &nbsp;<br />
                    <asp:Label ID="lblMsg1" runat="server" ForeColor="Red"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
    </div>
    </form>
</body></html>

