<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PrinterFriendly.aspx.vb" Inherits="DxTool.PrinterFriendly" MasterPageFile="~/Site1.Master" %>

 <asp:Content ContentPlaceHolderID="head" ID="head1" runat="server">
 <link href="/documentStyleEmail.css" rel="stylesheet" type="text/css" />
    </asp:Content>
<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">
  <asp:Image ID="surveyImage" runat="server" ImageUrl="/images/feedback.png" AlternateText="The Campus Diagnostic" />
  <asp:Label ID="results" runat="server"></asp:Label>
  <asp:Literal ID="printPage" runat="server"><script language="javascript">window.print();</script></asp:Literal>
</asp:Content>
