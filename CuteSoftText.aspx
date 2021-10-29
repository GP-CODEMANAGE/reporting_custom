<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CuteSoftText.aspx.cs" Inherits="CuteSoftText" %>

<%--<%@ Register TagPrefix="CE" Namespace="CuteEditor" Assembly="CuteEditor" %>--%>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
       <%-- <div>
            <CE:Editor ID="Editor1" runat="server" />
        </div>--%>
                 <tr>
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <div>
                            <%--<asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                <ContentTemplate>
                                    <cc:HtmlEditor ID="txtLetterText" runat="server" Height="200px" Width="400px" />
                                </ContentTemplate>
                            </asp:UpdatePanel>--%>
                            <FCKeditorV2:FCKeditor ID="txtLetterText" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                BasePath="~/FCKeditor/" Width="60%" Height="250px">
                            </FCKeditorV2:FCKeditor>
                            &nbsp;&nbsp;&nbsp;
                        </div>
                    </td>
                </tr>


                 <tr>
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <div>
                            <%--<asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                <ContentTemplate>
                                    <cc:HtmlEditor ID="txtLetterText" runat="server" Height="200px" Width="400px" />
                                </ContentTemplate>
                            </asp:UpdatePanel>--%>
                            <FCKeditorV2:FCKeditor ID="FCKeditor1" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                BasePath="~/FCKeditor/" Width="60%" Height="250px">
                            </FCKeditorV2:FCKeditor>
                            &nbsp;&nbsp;&nbsp;
                        </div>
                    </td>
                </tr>


                 <tr>
                    <td style="width: 25%"></td>
                    <td style="width: 60%">
                        <div>
                            <%--<asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="Conditional">
                                <ContentTemplate>
                                    <cc:HtmlEditor ID="txtLetterText" runat="server" Height="200px" Width="400px" />
                                </ContentTemplate>
                            </asp:UpdatePanel>--%>
                            <FCKeditorV2:FCKeditor ID="FCKeditor2" runat="server" ToolbarSet="Minimal" SkinPath="skins/silver/"
                                BasePath="~/FCKeditor/" Width="60%" Height="250px">
                            </FCKeditorV2:FCKeditor>
                            &nbsp;&nbsp;&nbsp;
                        </div>
                    </td>
                </tr>



        <asp:Button ID="Button1" runat="server" Text="Button" OnClick="Button1_Click" />
    </form>
</body>
</html>
