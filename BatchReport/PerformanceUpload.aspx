<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PerformanceUpload.aspx.cs"
    Inherits="BatchReport_BenchMarkUpload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Performance Upload</title>

    <script language="javascript" type="text/javascript">
    
     function CheckExtension()
    {
        var fup = document.getElementById('FileUpload1');
        var fileName = fup.value;
        var ext = fileName.substring(fileName.lastIndexOf('.') + 1);

        if(fileName == "")
        {
          alert("Please select file to upload");
          return false;
        }        

        if(ext != "xls")
        {
            alert("Please select '.xls' files only.");
            return false;
        }
     
    }
    
    function ClearLabel()
    {
        document.getElementById("lblError").innerHTML = "";
        if(document.getElementById("lnkIssues")!=null)
            document.getElementById("lnkIssues").style.display = "none";
        if(document.getElementById("lnkDuplicate")!=null)
            document.getElementById("lnkDuplicate").style.display = "none";
        if(document.getElementById("lnkMissing")!=null)
            document.getElementById("lnkMissing").style.display = "none";
        if (document.getElementById("lnkNotLoaded") != null)
            document.getElementById("lnkNotLoaded").style.display = "none";
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
                                <td colspan="3" class="Titlebig">
                                    Gresham Partners, LLC
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 18px" valign="top" colspan="3">
                                    <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    Sample File :</td>
                                <td style="width: 80%">
                                    <asp:LinkButton ID="lnkSamplefile" runat="server" OnClick="lnkSamplefile_Click">Download Sample File</asp:LinkButton>&nbsp;
                                </td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                </td>
                                <td style="width: 80%">
                                    <asp:LinkButton ID="lnkIssues" runat="server" OnClick="lnkIssues_Click" Visible="False">Perf records not inserted/updated</asp:LinkButton></td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    &nbsp;</td>
                                <td style="width: 80%">
                                    <asp:LinkButton ID="lnkDuplicate" runat="server" Visible="False" OnClick="lnkDuplicate_Click">Duplicate Perf records</asp:LinkButton></td>
                                <td style="width: 4px">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td style="width: 20%">
                                    &nbsp;</td>
                                <td style="width: 80%">
                                    <asp:LinkButton ID="lnkMissing" runat="server" Visible="False" OnClick="lnkMissing_Click">Missing Perf records</asp:LinkButton></td>
                                <td style="width: 4px">
                                    &nbsp;</td>
                            </tr>
                                                        <tr>
                                <td style="width: 20%">
                                </td>
                                <td style="width: 80%">
                                    <asp:LinkButton ID="lnkNotLoaded" runat="server" OnClick="lnkNotLoaded_Click" Visible="False">Perf records not Loaded</asp:LinkButton></td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr id="trBrowsefiles" runat="server">
                                <td style="width: 20%">
                                    <asp:Label ID="lblUploadFile" runat="server" Text="Upload File :"></asp:Label></td>
                                <td>
                                    <asp:FileUpload ID="FileUpload1" runat="server" onclick="ClearLabel();" Width="655px" /></td>
                                <td style="width: 4px">
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 59px">
                                </td>
                                <td valign="top" style="height: 59px">
                                    <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="Button1_Click1" OnClientClick="return CheckExtension();return false;" />&nbsp;</td>
                                <td style="width: 4px; height: 59px;">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
