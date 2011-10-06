<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="message.aspx.vb" Inherits="DxTool.message"  MasterPageFile="~/Site1.Master"%>
<asp:Content ID="LoginHeader" ContentPlaceHolderID="head" runat="server">
<link href="/documentStyle.css" rel="stylesheet" type="text/css" />
<link href="/documentLoginStyle.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">
<div id="regPage">
  <a href="http://www.outsidetheclassroom.com" target="_blank"><asp:image ID="loginBanner" runat="server" ImageUrl="/images/loginBanner.png" AlternateText="The Campus Diagnostic" /></a>
<div class="teal">Thanks for taking our newly-developed online diagnostic evaluation.</div> A few things to bear in mind before you get started: 
<ol style="float:left;">
<li><b>This is a partial assessment.</b>  The questions in this evaluation have been excerpted from a more comprehensive Diagnostic Inventory tool. This brief version evaluates only a limited scope of your prevention efforts; therefore, its measurement will be a less-than-perfect reflection of your campus’s efforts. </li>

<li><b>A note on benchmarking.</b>  When a critical mass of respondents has completed the Campus Diagnostic, we will share findings with all participants that will include benchmarks for national aggregates.  Your name and your institution’s name will never be identifiable in these reports—all data will be “blinded” for these purposes.</li>
<li><b>Be prepared for honest feedback.</b>  Once you complete the Campus Diagnostic, you will receive feedback and a grade.  We understand that receiving a letter grade alongside “suggestions for improvement” may be difficult.  Don’t be alarmed! Having this impartial assessment can be an important step towards improving your campus’s prevention efforts.</li>
    <asp:ImageButton  ID="btnStartImg" style="float:right; padding:15px;" runat="server" ImageUrl="/images/GetStartedBtn.png" AlternateText="Get Started" />

<li><b>We want your feedback!  </b>   As we continue to improve this tool, we welcome your suggestions, comments, and questions. Please share your thoughts at campusdiagnostic@outsidetheclassroom.com.
</li>
</ol>


</div>

</asp:Content>
<asp:Content id="ContentF" ContentPlaceHolderID="footer" runat="server"><div  class="footer">
Please refer to our <a href="http://www.outsidetheclassroom.com/terms-and-conditions.aspx" target="_new">Terms and Conditions</a> before taking The Campus Diagnostic.
</div>
</asp:Content>

