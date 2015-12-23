<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="GroupDocs.ReportGenerator.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>:: Next Generation GroupDocs Report Generator</title>
    <link href="/Css/ReportGenerator.css"
        rel="stylesheet" type="text/css" media="all" />
    <script type="text/javascript">
        function sourceChanged() {
            var objSource = document.getElementById("<%=ddlSource.ClientID%>");
            var objTables = document.getElementById("<%=ddlTables.ClientID%>");
            var objViews = document.getElementById("<%=ddlViews.ClientID%>");
            var objQuery = document.getElementById("<%=txtCustomQuery.ClientID%>");

            objTables.style["display"] = "none";
            objViews.style["display"] = "none";
            objQuery.style["display"] = "none";

            switch (objSource.value) {
                case "0":
                    objTables.style["display"] = "block";
                    break;
                case "1":
                    objViews.style["display"] = "block";
                    break;
                case "2":
                    objQuery.style["display"] = "block";
                    break;
            }
        }
    </script>

</head>
<body>
    <form id="form1" runat="server">

        <div class="GroupDocsReportGenerator">

            <div>
                <asp:Label ID="lblMessage" runat="server" Font-Bold="true" Text="" ForeColor="Maroon"></asp:Label>
            </div>
            <table width="100%">
                <tr>
                    <td align="center">
                        <table class="tblBack">
                            <tr>
                                <td colspan="2">
                                    <h2>GroupDocs: Report Generator V 1.0 using GroupDocs.Assembly for .NET</h2>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">&nbsp;
                                </td>
                            </tr>
                              <tr>
                                <td>&nbsp;Template File (MS Word, MS Excel):
                                </td>
                                <td>
                                    <asp:FileUpload ID="fuTemplate" runat="server" CssClass="input" Width="485px" /><asp:Button ID="UploadTemplate" CssClass="buttonClass" runat="server" Visible="true"
                                        OnClick="UploadTemplate_Click" Text="Upload Template"></asp:Button>

                                </td>
                            </tr>
                            <tr>
                                <td>&nbsp;Connection String <i>(MS SQL Server)</i>:
                                </td>
                                <td>
                                    <asp:TextBox ID="txtConString" runat="server" Text="SERVER=(local); DATABASE=sampledb; USER ID=sa; PASSWORD=abc123;" MaxLength="200" CssClass="input"
                                        Width="485px"></asp:TextBox>
                                    <asp:Button ID="TestConnectionAndLoad" CssClass="buttonClass" runat="server" Visible="true"
                                        OnClick="TestConnectionAndLoad_Click" Text="Test Connection"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td>&nbsp;Type of Data Source:
                                </td>
                                <td>
                                    <asp:DropDownList runat="server" ID="ddlSource" CssClass="input" Width="99%">
                                        <asp:ListItem Selected="True" Value="0" Text="Tables"></asp:ListItem>
                                        <asp:ListItem Value="1" Text="Views"></asp:ListItem>
                                        <asp:ListItem Value="2" Text="Custom SQL Query"></asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td>&nbsp;Data Source:
                                </td>
                                <td>
                                    <asp:DropDownList runat="server" ID="ddlTables" CssClass="input" Width="99%">
                                    </asp:DropDownList>
                                    <asp:DropDownList runat="server" ID="ddlViews" Style="display: none;" CssClass="input"
                                        Width="600px">
                                    </asp:DropDownList>
                                    <asp:TextBox ID="txtCustomQuery" runat="server" TextMode="MultiLine" CssClass="input"
                                        Style="display: none;" MaxLength="1000" Width="97%" Height="120px"></asp:TextBox>
                                </td>
                            </tr>
                          
                            <tr>
                                <td>&nbsp;
                                </td>
                                <td>
                                    <asp:Button ID="GenerateReport" CssClass="buttonClass" runat="server" Visible="true"
                                        OnClick="GenerateReport_Click" Text="Generate Report"></asp:Button>&nbsp;&nbsp;
                                    <asp:Button ID="ClearForm" CssClass="buttonClass" runat="server" Visible="true"
                                        OnClick="ClearForm_Click" Text="Clear"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">&nbsp;
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
