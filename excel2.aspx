<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="excel2.aspx.cs" Inherits="excel.excel2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            margin-left: 181px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
       Import Excel File:    
        <asp:FileUpload ID="FileUploadToServer" width="300px" runat="server" />    
        <br />    
        <br />    
        <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Upload File" ValidationGroup="vg" Style="width:99px" />    
        <br />    
        <br />    
        <asp:Label ID="lblMsg" runat="server" ForeColor="Green" Text=""></asp:Label>    
        <br />    
              <asp:GridView ID="GridView1" runat="server" CellPadding="4"  EmptyDataText="No Records Found!" Height="25px" ForeColor="#333333" GridLines="None" CssClass="auto-style1" Width="295px">    
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />    
            <EditRowStyle BackColor="#999999" />    
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />    
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />    
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />    
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />    
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />    
               
        </asp:GridView>   
<%--        <asp:GridView ID="GridView1" runat="server"  GridLines="None" EmptyDataText="No Records Found!" Height="25px">    
            <RowStyle Width="175px" />
            <EmptyDataRowStyle BackColor="Silver" BorderColor="#999999" BorderStyle="solid" borderwidth="1px" ForeColor="#003300"/>
            <HeaderStyle BackColor="#6699ff" BorderColor="#333333" BorderStyle="Solid" BorderWidth="1px" VerticalAlign="Top" Width="200px" Wrap="true" />
        </asp:GridView>--%>    
        
    </div>    
    </form>    
</body>    
</html> 