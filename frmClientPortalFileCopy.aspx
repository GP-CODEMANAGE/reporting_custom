<%@ Page Language="C#" AutoEventWireup="true" CodeFile="frmClientPortalFileCopy.aspx.cs"
    Inherits="frmClientPortalFileCopy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <link id="Link1" href="./common/gresham.css" rel="stylesheet" type="text/css" />
    <link href="./common/Calendar.css" rel="stylesheet" type="text/css" />
    <script src="./common/Calendar.js" type="text/javascript"></script>
    <%--<script src="ckeditor/jquery-3.1.1.js"></script>--%>
    <script src="ckeditor/jscript_1.12.4.js" type="text/javascript"></script>
    <title>Client Portal File Copy</title>
    <%--<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.12.4.min.js">
      
        }

    </script>--%>
    <script type="text/javascript">
        function callme() {

            var filePath = document.getElementById('<%=FileUpload1.ClientID %>').value;
            //var filePath = document.getElementById("FileUpload1").value;
            //function getNameFromPath(filePath) 
         
            document.getElementById("txtFilePath").value = filePath;
         

            var objRE = new RegExp(/([^\/\\]+)$/);
            // debugger;
            var strName = objRE.exec(filePath);

            if (strName == null) {
                return null;
            }
            else {
                //return strName[0];
              //  document.getElementById("txtfilename").value = strName[0];
                document.getElementById("FileName").value = strName[0];
                document.getElementById("Button1").click();
                
            }
        }

        function onFolderPathClick(Path) {
            //$("#txtFolderPath").attr("value", Path);
            document.getElementById("txtFolderPath").value = Path;
        }
       
        $(function () {
            $(document).click('change', 'FileUpload1', function () {
                alert('bye');
                var filePath = $("#uploadFile").attr('value');
                var fileName = filePath.split("\\")[filePath.split("\\").length - 1];
                $("#TextBox1").attr("value", fileName);
            })
        })

        $("#FileUpload1").click('change', function () {
            var filePath = $("#uploadFile").attr('value');
            var fileName = filePath.split("\\")[filePath.split("\\").length - 1];
            $("#TextBox1").attr("value", fileName);
        })


        //function getYearDropDown() {
        //    //$("#txtFolderPath").attr("value", Path);
        //    var value = ddlClientPortalPath.find("option:selected").text();

        //}

    </script>
    <script type="text/javascript" language="javascript">
        function CheckAllEmp(Checkbox) {
            var gvList = document.getElementById("<%=gvList.ClientID %>");
            for (i = 1; i < gvList.rows.length; i++) {
                gvList.rows[i].cells[0].getElementsByTagName("INPUT")[0].checked = Checkbox.checked;
            }
        }
    </script>
    <script type="text/javascript">
        var TotalChkBx;
        var Counter;

        window.onload = function () {
            //Get total no. of CheckBoxes in side the GridView.
            TotalChkBx = parseInt('<%= this.gvList.Rows.Count %>');

            //Get total no. of checked CheckBoxes in side the GridView.
            Counter = 0;
        }

        function HeaderClick(CheckBox) {
            //Get target base & child control.
            var TargetBaseControl =
       document.getElementById('<%= this.gvList.ClientID %>');
            var TargetChildControl = "chkbSelectBatch";

            //Get all the control of the type INPUT in the base control.
            var Inputs = TargetBaseControl.getElementsByTagName("input");

            //Checked/Unchecked all the checkBoxes in side the GridView.
            for (var n = 0; n < Inputs.length; ++n)
                if (Inputs[n].type == 'checkbox' &&
                Inputs[n].id.indexOf(TargetChildControl, 0) >= 0)
                    Inputs[n].checked = CheckBox.checked;

            //Reset Counter
            Counter = CheckBox.checked ? TotalChkBx : 0;
        }

        function ChildClick(CheckBox, HCheckBox) {
            //get target control.
            var HeaderCheckBox = document.getElementById(HCheckBox);

            //Modifiy Counter; 
            if (CheckBox.checked && Counter < TotalChkBx)
                Counter++;
            else if (Counter > 0)
                Counter--;

            //Change state of the header CheckBox.
            if (Counter < TotalChkBx)
                HeaderCheckBox.checked = false;
            else if (Counter == TotalChkBx)
                HeaderCheckBox.checked = true;
        }

    </script>

     <script type="text/javascript">
         function do_totals1() {
             document.all.pleasewaitScreen.style.visibility = "visible";
             window.setTimeout('do_totals2()', 1)
         }
         function do_totals2() {
             calc_totals();
             document.all.pleasewaitScreen.style.visibility = "hidden";
         }
  </script>

    <%--<script type="text/javascript">
        function callme(oFile) {
            document.getElementById("txtFilename2").value = oFile.value;
        }

</script>--%><%-- <script>
        function getNameFromPath(FileUpload1) {
            var objRE = new RegExp(/([^\/\\]+)$/);
            var strName = objRE.exec(FileUpload1);

            if (strName == null) {
                return null;
            }
            else {
                return strName[0];
            }
        }
    </script>--%>
    <style type="text/css">
        .Titlebig
        {
            font-family: Frutiger 55 Roman;
            font-size: 14pt;
            font-weight: normal;
            text-decoration: none;
        }
        
        span
        {
            font-family: Frutiger 55 Roman;
            font-size: 12pt;
        }
        
        .style4
        {
            width: 1%;
        }
        
        .auto-style4
        {
            width: 11%;
        }
        
        .auto-style5
        {
            width: 26%;
        }
        
        .auto-style6
        {
            width: 10%;
        }
        
        .gridview
        {
            border-style: none !important;
        }
        
        .textboxwidth
        {
            width: 337px;
        }
    </style>

    <style type="text/css">
    .modal
    {
        position: fixed;
        top: 0;
        left: 0;
        background-color: black;
        z-index: 99;
        opacity: 0.8;
        filter: alpha(opacity=80);
        -moz-opacity: 0.8;
        min-height: 100%;
        width: 100%;
    }
    .loading
    {
        font-family: Arial;
        font-size: 10pt;
        border: 5px solid #67CFF5;
        width: 200px;
        height: 100px;
        display: none;
        position: fixed;
        background-color: White;
        z-index: 999;
    }
</style>
</head>
<body>

 

    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
        <table style="width: 100%">
            <tr>
                <td colspan="5">
                    <img src="images/Gresham_Logo__.jpg" />
                </td>
            </tr>
            <tr>
                <td colspan="3" class="Titlebig">
                    Gresham Partners, LLC
                </td>
            </tr>
            <tr>
                <td class="Titlebig" colspan="3">
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Client Portal File Copy
                </td>
            </tr>
            <tr>
                <td style="height: 18px" valign="top" colspan="3">
                    <br />
                    <asp:Label ID="lblMessage" runat="server" ForeColor="Red"></asp:Label>
                    <br />
                </td>
            </tr>
            <tr>
                <%-- <td style="height: 18px" valign="top" colspan="3">
                        <asp:Label ID="lblSelect" runat="server"
                            Text="Please select the following fields to generate your report: "></asp:Label>
                        <br />
                    </td>--%>
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;

                    
            <asp:Label ID="lablmsg" runat="server" Text="" ForeColor="Green"></asp:Label>

                </td>
                <td class="style4" style="color: #FFFFFF">
                    :
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    <asp:Label ID="lblSO" runat="server" Text="Select a file to be uploaded"></asp:Label>
                </td>
                <td class="style4" style="color: #FFFFFF">
                    
                    :
                </td>
                <td colspan="2">
                   
               <asp:FileUpload ID="FileUpload1" runat="server" Width="96px" onchange="callme(this)" AutoPostBack="true" 
                                Text="Browse" Height="24px"   />
                 
                    <asp:TextBox ID="txtFilePath" runat="server" Width="545px"></asp:TextBox>
                         
                </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    <asp:Label ID="lblHH" runat="server" Text="Change a File Name to"></asp:Label>
                </td>
                <td class="style4" style="color: #FFFFFF">
                    :
                </td>
                <td colspan="2">
                    <asp:TextBox ID="txtfilename" runat="server" Width="372px" Visible="false"></asp:TextBox>
                    <asp:TextBox ID="FileName" runat="server" Width="344px"></asp:TextBox>
                    <%--   OnTextChanged="txtfilename_TextChanged"--%>
                </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    <asp:Label ID="Label1" runat="server" Text="Document Type Tags"></asp:Label>
                </td>
                <td class="style4" style="color: #FFFFFF">
                    :
                </td>
                <td colspan="2">
                    <%--  <asp:TextBox ID="txtFolderPath" runat="server" AutoPostBack="true"  BackColor="#C7EDF7"
                            Width="431px"  ></asp:TextBox>--%>
                    <asp:TextBox ID="txtFolderPath" runat="server" BackColor="#C7EDF7" Width="373px" Visible="false"></asp:TextBox>
                  <%--  <asp:DropDownList ID="ddlClientPortalPath" runat="server" Width="298px"  OnSelectedIndexChanged="ddlClientPortalPath_SelectedIndexChanged" ></asp:DropDownList> --%>
                    <asp:DropDownList ID="ddlPortalPath" runat="server"  OnSelectedIndexChanged="ddlPortalPath_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                     <asp:DropDownList ID="ddlYear" runat="server" Width="98px" Visible="false" ></asp:DropDownList> 
                    <%--   <input type="text" id="txtFolderPath" name="txtFolderPath" class="textboxwidth" />--%>
                &nbsp;
                  <%--  <asp:Label ID="Label6" runat="server" Text="Label"></asp:Label>--%>
                </td>
            </tr>
           <tr>
                    <td class="auto-style4">_Test Flag&nbsp;
                        
                    </td>
                    <td class="style4" style="color: #FFFFFF">: </td>
                    <td class="auto-style48">
                        <asp:CheckBox ID="cbTest" runat="server"  />
                    </td>

                </tr>
            <tr>
                <td>

                    &nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style6">
                    <asp:Button ID="btnUploadFile" runat="server" OnClick="btnUploadFile_Click" Text="Upload File to Client Portal"
                        Width="370px" />
                </td>
                <td class="auto-style47" style="color: #FFFFFF">
                </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;
                </td>
                <td class="style4">
                </td>
                <td class="auto-style4">
                    <asp:Label ID="Label2" runat="server" Text="Start Date"></asp:Label>
                    <asp:TextBox ID="txtStartDate" runat="server" AutoPostBack="true" Width="139px"></asp:TextBox>
                    <a onclick="showCalendarControl(txtStartDate,'restrict=true,close=true')">
                        <img id="img2" alt="" border="0" src="images/calander.png" /></a>
                </td>
                <td class="auto-style5" align="left">
                </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;
                </td>
                <%--<td class="auto-style4">
                        <asp:Button ID="Button1" runat="server"  Text="Upload File to Client Portal" />
                    </td>--%>
                <%--<td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:TextBox ID="txtAsOfDate0" runat="server" Width="119px" onChange="selectMonths(this.value);" OnTextChanged="txtAsOfDate_TextChanged" AutoPostBack="True"></asp:TextBox>
                        <a onclick="showCalendarControl(txtAsOfDate,'restrict=true,close=true')">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsOfDate" Display="Dynamic" ErrorMessage="Select As of Date" SetFocusOnError="True">*</asp:RequiredFieldValidator>
                    </td>--%>
                <td class="style4">
                </td>
                <td class="auto-style4">
                    <asp:Label ID="Label3" runat="server" Text="AsOfDate"></asp:Label>
                    <asp:TextBox ID="txtAsOfDate" runat="server" AutoPostBack="true" Width="138px" OnTextChanged="txtAsOfDate_TextChanged1"></asp:TextBox>
                    <a onclick="showCalendarControl(txtAsOfDate,'restrict=true,close=true')">
                        <img id="img1" alt="" border="0" src="images/calander.png" /></a>
                </td>





                <td class="auto-style5" align="left">
                </td>
            </tr>
            <tr>
                <td>

                </td>
                <td>

                </td>
<td>
  <asp:Label ID="Label6" runat="server" Text="HouseholdType-"></asp:Label>
    <asp:DropDownList ID="ddlHouseHoldType" runat="server" AutoPostBack="true" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" >
                                </asp:DropDownList>
</td><td>

     </td>
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;
                </td>
                <%--<td class="auto-style4">
                        <asp:Button ID="Button1" runat="server"  Text="Upload File to Client Portal" />
                    </td>--%>
                <%--<td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:TextBox ID="txtAsOfDate0" runat="server" Width="119px" onChange="selectMonths(this.value);" OnTextChanged="txtAsOfDate_TextChanged" AutoPostBack="True"></asp:TextBox>
                        <a onclick="showCalendarControl(txtAsOfDate,'restrict=true,close=true')">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsOfDate" Display="Dynamic" ErrorMessage="Select As of Date" SetFocusOnError="True">*</asp:RequiredFieldValidator>
                    </td>--%>
                <td class="style4">
                </td>
                <td class="auto-style4">
                    <asp:Label ID="Label4" runat="server" Text="Fund-"></asp:Label>
                    <asp:DropDownList ID="ddlFund" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlFund_SelectedIndexChanged" Height="18px">
                    </asp:DropDownList>
                </td>
                <td class="auto-style5" align="left">
            </tr>
            <tr>
                <td class="auto-style6">
                    &nbsp;
                </td>
                <%--<td class="auto-style4">
                        <asp:Button ID="Button1" runat="server"  Text="Upload File to Client Portal" />
                    </td>--%>
                <%--<td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:TextBox ID="txtAsOfDate0" runat="server" Width="119px" onChange="selectMonths(this.value);" OnTextChanged="txtAsOfDate_TextChanged" AutoPostBack="True"></asp:TextBox>
                        <a onclick="showCalendarControl(txtAsOfDate,'restrict=true,close=true')">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsOfDate" Display="Dynamic" ErrorMessage="Select As of Date" SetFocusOnError="True">*</asp:RequiredFieldValidator>
                    </td>--%>
                <td class="style4">
                </td>
                <td class="auto-style4">
                    <asp:Label ID="Label5" runat="server" Text="Select the list of clients"></asp:Label>

                   
                </td>
                <td class="auto-style5" align="left">
                </td>
            </tr>
            <tr>
                <%--<td class="auto-style4">
                        <asp:Button ID="Button1" runat="server"  Text="Upload File to Client Portal" />
                    </td>--%>
                <%--<td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:TextBox ID="txtAsOfDate0" runat="server" Width="119px" onChange="selectMonths(this.value);" OnTextChanged="txtAsOfDate_TextChanged" AutoPostBack="True"></asp:TextBox>
                        <a onclick="showCalendarControl(txtAsOfDate,'restrict=true,close=true')">
                            <img id="img1" alt="" border="0" src="images/calander.png" /></a><asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAsOfDate" Display="Dynamic" ErrorMessage="Select As of Date" SetFocusOnError="True">*</asp:RequiredFieldValidator>
                    </td>--%>
                <td align="left" valign="top" colspan="2">
                    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" BorderStyle="None"
                        CssClass="gridview" GridLines="None">
                        <Columns>
                            <asp:TemplateField HeaderText="" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <%--<asp:HyperLink ID="link1" runat="server" Text='<%# Eval("Rootfolder") %>'  onclick="returm onFolderPathClick(<%#Eval("path") %>)"/>--%>
                                    <a href="javascript:;" onclick='return onFolderPathClick("<%#Eval("DocumentType") %>");'>
                                        <%#Eval("DocumentType") %>
                                    </a>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <%--   <asp:HyperLink ID="link2" runat="server" Text='<%# Eval("SubFolder") %>' />--%>
                                    <a href="javascript:;" onclick='return onFolderPathClick("<%#Eval("Year") %>");'>
                                        <%#Eval("Year") %>
                                    </a>
                                </ItemTemplate>
                            </asp:TemplateField>
                           <%-- <asp:TemplateField HeaderText="" ItemStyle-HorizontalAlign="Left" FooterStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <%--  <asp:HyperLink ID="link3" runat="server" Text='<%# Eval("subSubfolder") %>' />
                                    <a href="javascript:;" onclick='return onFolderPathClick("<%#Eval("Year") %>");'>
                                        <%#Eval("Year") %>
                                    </a>
                                </ItemTemplate>
                            </asp:TemplateField>--%>
                            <%-- <asp:BoundField DataField="path" HeaderText="path" />--%>
                        </Columns>
                    </asp:GridView>
                </td>
                <td align="left">
                    <br />
                    <contenttemplate>
                            <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="false" TabIndex="1"
                                ToolTip="Batch List" Width="49%" GridLines="None">
                                <Columns>
                                    <asp:BoundField DataField="Ssi_batchId" HeaderText="Ssi_batchId" Visible="False" />
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkboxSelectAll" runat="server" Checked="false" onclick="HeaderClick(this);" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkbSelectBatch" runat="server" Checked="false" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="BatchName" HeaderText="All clients" SortExpression="BatchName"
                                        ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" />
                                </Columns>
                                <HeaderStyle Height="10px" />
                            </asp:GridView>
                </td>
                <td class="auto-style5">
                    &nbsp;&nbsp;&nbsp;  
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Button ID="Button1" runat="server" Text="" OnClick="Button1_Click" Width="16px"  BackColor="#F0F0F0" BorderStyle="None" />
                </td>
            </tr>
        </table>
    </div>
        <DIV id="pleasewaitScreen" style="Z-INDEX: 5; LEFT: 35%; VISIBILITY: hidden; POSITION: absolute; TOP: 40%; width: 256px; height: 191px;">
    <TABLE  border="1" style="width: 256px; height: 191px">
     <TR>
	<TD vAlign="middle" align="center" width="100%" bgColor="#ffffff" height="100%"><BR>
	 <BR>
	  <IMG src="Images/ajax-loader.gif" align="middle"><%-- <FONT face="Lucida Grande, Verdana, Arial, sans-serif" color="#000066" size="5">
           <B> Please wait...</FONT>--%>
	 <BR>
	 <BR>
        </TD>
    </TR>
   </TABLE>
</DIV>
    </form>
  
</body>
</html>
