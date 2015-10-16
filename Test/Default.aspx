<%@ Page Title="主页" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="Test._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <asp:Button ID="btnDt" runat="server" Text="dt" onclick="btnDt_Click" />
    <asp:Button ID="btnDs" runat="server" Text="ds" onclick="btnDs_Click" />
    <asp:Button ID="btnMany" runat="server" Text="多表头" onclick="btnMany_Click" />
</asp:Content>
