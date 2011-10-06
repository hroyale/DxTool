<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="moreinfo.aspx.vb" Inherits="DxTool.moreinfo" MasterPageFile="~/Site1.Master"%>


<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">
    <a href="http://www.outsidetheclassroom.com" target="_blank"><asp:Image ID="surveyImage" runat="server" ImageUrl="/images/feedback.png" AlternateText="The Campus Diagnostic" /></a>
<div id="return">
<asp:ImageButton ID="backBtn" runat="server" ImageUrl="/images/back.png" AlternateText="Return to Grades" />
</div>
<div class="infopage">
<div class="teal">More information on the full Diagnostic Inventory:</div>
<div>
The Campus Diagnostic is comprised of a sample of questions excerpted from two sections of a 
much more extensive instrument, The Alcohol Prevention Coalition Diagnostic Inventory. The longer tool was
developed by Outside The Classroom to comprehensively assesses many dimensions of campus alcohol prevention at 
our <a href="http://www.outsidetheclassroom.com/solutions/higher-education/Coalition.aspx">Alcohol Prevention Coalition</a> partner institutions. These dimensions extend beyond a close examination of 
prevention programming and the degree of institutional support for alcohol prevention on campus, scrutinizing 
campus alcohol policies, their enforcement and adjudication, adherence to processes deemed critical to success 
in alcohol prevention, and the extent of relationships with a variety of constituencies that are important to 
prevention success. Completion of this full Diagnostic Inventory allows Alcohol Prevention Coalition campuses to not 
only pinpoint areas of strength and weakness and set goals for improvement, but also to benchmark their alcohol 
prevention progress annually and see how they compare to other campuses that have taken the instrument. 
</div>
<div class="bottom">
<asp:ImageButton ID="consultBtn" runat="server" ImageUrl="/images/consult.png" AlternateText="Would you like to schedule a consultation to learn more insights" />
</div></div>
</asp:Content>
