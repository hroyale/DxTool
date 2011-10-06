<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EmailProgram.aspx.vb" Inherits="DxTool.EmailProgram" MasterPageFile="~/Site1.Master"%>


<asp:Content ID="LoginHeader" ContentPlaceHolderID="head" runat="server">
<link href="/documentStyle.css" rel="stylesheet" type="text/css" />
<link href="/documentLoginStyle.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">




<div class="faqContainer">

   
    <asp:Panel ID="showForm" runat="server" Visible="true"> <div>
    <asp:Label ID="errorMsg" runat="server" Visible="false"></asp:Label>
    <table>
    <tr>
    <td class="title"><b>To:</b></td>
    <td><asp:TextBox ID="toEmail" runat="server" Columns="40" BackColor="Yellow"> enter your friend's email address here </asp:TextBox></td>
    </tr>
    <tr>
    <td class="title"><b>From:</b></td>
    <td><asp:TextBox ID="fromEmail" runat="server" Columns="40" BackColor="Yellow"> enter your email address here </asp:TextBox>
      </td>
    </tr>
    <tr>
    <td class="title"><b>Subject:</b></td>
    <td><asp:TextBox ID="subject" runat="server" Columns="70">From Your Friend: Please Take The Campus Diagnostic</asp:TextBox></td>
    </tr>
    <tr>
    <td colspan="2">
    <br /><br />
    <asp:TextBox ID="emailBody1" runat="server" Rows="14" Columns="70" Wrap=true TextMode=MultiLine>
 </asp:textBox>
 <div> Here&#8217;s the link: <asp:hyperlink id="parentURL" runat="server"></asp:hyperlink>.   
 <br /><br />
 </div>
  
<br /> 
 <asp:TextBox ID="emailBody2" runat="server" Rows="5" Columns="70" Wrap=true TextMode=MultiLine>

Signed,
Your Friend
    </asp:TextBox>
    
    
    </td>
    
    </tr>
    </table>
        <asp:Button ID="btnSend" runat="server" Text="Send Email" />

    </div>
</asp:Panel>
<asp:Panel ID="showMessage" runat="server" Visible="false">
<h2>Thank you!</h2>
    The Email has been sent. You may close this window to return to your grades.
<br /><br />

</asp:Panel>
						

</asp:Content>

