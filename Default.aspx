<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="DxTool._Default" MasterPageFile="~/Site1.Master" %>
<asp:Content ID="LoginHeader" ContentPlaceHolderID="head" runat="server">
<link href="/documentStyle.css" rel="stylesheet" type="text/css" />
<link href="/documentLoginStyle.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">
<div id="regPage">
  <a href="http://www.outsidetheclassroom.com" target="_blank"><asp:image ID="loginBanner" runat="server" ImageUrl="/images/loginBanner.png" AlternateText="The Campus Diagnostic" /></a>
The following questionnaire provides an evaluation of your campus’s alcohol prevention programming and the degree to which prevention is prioritized at your institution.  After completing the questionnaire, you will immediately receive a grade and qualitative feedback regarding where your institution stands with respect to alcohol prevention.  
<div class="teal">This questionnaire takes approximately 5 minutes to complete. </div>
<div id="rightSide">
 <div class="error"><asp:Label ID="errorMsg" runat="server" Visible="false"></asp:Label></div>
<asp:ImageButton ID="btnSubmitImg" class="submitBtn" runat="server" ImageUrl="/images/Continue.png" AlternateText="submit" />
</div>
<asp:Button ID="btnSubmit" runat="server" Text="Continue" class="submit" visible="false"/>
<div id="register">
<div>First Name: &nbsp; <asp:TextBox ID="fName" runat="server" Columns="20"></asp:TextBox></div>
<div> Last Name: &nbsp; <asp:TextBox ID="lName" runat="server" Columns="20"></asp:TextBox></div>
<div>Title: &nbsp; <asp:TextBox ID="title" runat="server" Columns="20"></asp:TextBox></div>
<div>Institution: &nbsp; <asp:TextBox id="school" runat="server" Columns="20"></asp:TextBox> in&nbsp; <asp:DropDownList id="State" runat="server">
    <asp:ListItem Value="oo">Select State</asp:ListItem>
    <asp:ListItem Value="AL">Alabama</asp:ListItem>
    <asp:ListItem Value="AK">Alaska</asp:ListItem>
    <asp:ListItem Value="AZ">Arizona</asp:ListItem>
    <asp:ListItem Value="AR">Arkansas</asp:ListItem>
    <asp:ListItem Value="CA">California</asp:ListItem>
    <asp:ListItem Value="CO">Colorado</asp:ListItem>
    <asp:ListItem Value="CT">Connecticut</asp:ListItem>
    <asp:ListItem Value="DC">District of Columbia</asp:ListItem>
    <asp:ListItem Value="DE">Delaware</asp:ListItem>
    <asp:ListItem Value="FL">Florida</asp:ListItem>
    <asp:ListItem Value="GA">Georgia</asp:ListItem>
    <asp:ListItem Value="HI">Hawaii</asp:ListItem>
    <asp:ListItem Value="ID">Idaho</asp:ListItem>
    <asp:ListItem Value="IL">Illinois</asp:ListItem>
    <asp:ListItem Value="IN">Indiana</asp:ListItem>
    <asp:ListItem Value="IA">Iowa</asp:ListItem>
    <asp:ListItem Value="KS">Kansas</asp:ListItem>
    <asp:ListItem Value="KY">Kentucky</asp:ListItem>
    <asp:ListItem Value="LA">Louisiana</asp:ListItem>
    <asp:ListItem Value="ME">Maine</asp:ListItem>
    <asp:ListItem Value="MD">Maryland</asp:ListItem>
    <asp:ListItem Value="MA">Massachusetts</asp:ListItem>
    <asp:ListItem Value="MI">Michigan</asp:ListItem>
    <asp:ListItem Value="MN">Minnesota</asp:ListItem>
    <asp:ListItem Value="MS">Mississippi</asp:ListItem>
    <asp:ListItem Value="MO">Missouri</asp:ListItem>
    <asp:ListItem Value="MT">Montana</asp:ListItem>
    <asp:ListItem Value="NE">Nebraska</asp:ListItem>
    <asp:ListItem Value="NV">Nevada</asp:ListItem>
    <asp:ListItem Value="NH">New Hampshire</asp:ListItem>
    <asp:ListItem Value="NJ">New Jersey</asp:ListItem>
    <asp:ListItem Value="NM">New Mexico</asp:ListItem>
    <asp:ListItem Value="NY">New York</asp:ListItem>
    <asp:ListItem Value="NC">North Carolina</asp:ListItem>
    <asp:ListItem Value="ND">North Dakota</asp:ListItem>
    <asp:ListItem Value="OH">Ohio</asp:ListItem>
    <asp:ListItem Value="OK">Oklahoma</asp:ListItem>
    <asp:ListItem Value="OR">Oregon</asp:ListItem>
    <asp:ListItem Value="PA">Pennsylvania</asp:ListItem>
    <asp:ListItem Value="RI">Rhode Island</asp:ListItem>
    <asp:ListItem Value="SC">South Carolina</asp:ListItem>
    <asp:ListItem Value="SD">South Dakota</asp:ListItem>
    <asp:ListItem Value="TN">Tennessee</asp:ListItem>
    <asp:ListItem Value="TX">Texas</asp:ListItem>
    <asp:ListItem Value="UT">Utah</asp:ListItem>
    <asp:ListItem Value="VT">Vermont</asp:ListItem>
    <asp:ListItem Value="VA">Virginia</asp:ListItem>
    <asp:ListItem Value="WA">Washington</asp:ListItem>
    <asp:ListItem Value="WV">West Virginia</asp:ListItem>
    <asp:ListItem Value="WI">Wisconsin</asp:ListItem>
    <asp:ListItem Value="WY">Wyoming</asp:ListItem>
    <asp:ListItem Value="OT">- Other -</asp:ListItem>    
</asp:DropDownList> 
</div>
<div>Email Address: &nbsp; <asp:TextBox ID="email" runat="server" Columns="30"></asp:TextBox><br><em>Your email address is required in order to receive a customized report of your results</em></div>
<!--//<div style="display:none;">No. of Undergrad Students: &nbsp; <asp:TextBox ID="studentPop" runat="server" Columns="20"></asp:TextBox></div> //-->
</div>
</div>

</asp:Content>
<asp:Content id="ContentF" ContentPlaceHolderID="footer" runat="server"><div  class="footer">
Please refer to our <a href="http://www.outsidetheclassroom.com/terms-and-conditions.aspx" target="_new">Terms and Conditions</a> before taking The Campus Diagnostic.
</div>
</asp:Content>
