<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Feedback.aspx.vb" Inherits="DxTool.Feedback" MasterPageFile="~/Site1.Master" %>

 <asp:Content ContentPlaceHolderID="head" ID="head1" runat="server">
 <link href="/documentStyle.css" rel="stylesheet" type="text/css" />

 <script language="javascript">
     function openWindow(url) {
         window.open(url, "popup");
     }
 </script>
  </asp:Content>
<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server">
<div id="feedbackPage">
  <a href="http://www.outsidetheclassroom.com" target="_blank"><asp:Image ID="surveyImage" runat="server" ImageUrl="/images/feedback.png" AlternateText="The Campus Diagnostic" /></a>
<div id="FBRight">
<asp:ImageButton ID="printThisBtn" class="rtCol" runat="server" ImageUrl="/images/print.png" AlternateText="Print" />
<asp:ImageButton ID="consultBtn1" class="rtCol" runat="server" ImageUrl="/images/consult_side.png" AlternateText="Would you like to schedule a consultation to learn more insights" />
<asp:ImageButton ID="readMoreBtn" class="rtCol" runat="server" ImageUrl="/images/readmore.png" AlternateText="Read More about our full Diagnostic Inventory" />
<asp:HyperLink class="rtCol" ID="forwardLnk" runat="server" NavigateUrl="/EmailProgram.aspx" Target="_blank"><asp:Image ID="forwardBtn" class="rtCol" runat="server" ImageUrl="/images/forward.png" AlternateText="Forward The Campus Diagnostic to a Friend" /></asp:HyperLink>
<asp:HyperLink class="rtCol" ID="feedbackLnk" runat="server" NavigateUrl="http://www.surveymonkey.com/s/TZBFJBM " Target="_blank"><asp:Image ID="feedbackBtn" class="rtCol" runat="server" ImageUrl="/images/feedbackBtn.png" AlternateText="Tell us what you think." /></asp:HyperLink>

</div>  
<div id="FBHead" class="teal">Thank you for completing Outside The Classroom’s Campus Diagnostic.</div> 
<div id="FBIntro">Feedback on your campus’s alcohol prevention programming and the degree to which prevention is a priority at your institution is provided below:
</div>
 <asp:Panel ID="ProgramScore" runat="server" Visible="true">
 <div class="greenBar">Programmatic Grade:</div>
<div id="scoreCard">
 
<div class="grade"><asp:Label ID="LetterGrade" runat="server"></asp:Label></div>
<asp:Panel ID="ShowA2" runat="server" Visible="false">
<!--A+ Message<br /><br />-->
<asp:label id="PA2" runat="server">
Your alcohol prevention programming overall is excellent, with an optimal mix of <a href="calculate.aspx#def" target="new">universal, selective, and indicated</a> programming elements. The prevention programs you are using have a strong basis in the research literature. We encourage you to collect data to evaluate the success of your campus programs and then disseminate your results.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowA" runat="server" Visible="false">
<!--A Message<br /><br />-->
<asp:label id="PA" runat="server">
Your alcohol prevention programming overall is strong, with a strong mix of <a href="calculate.aspx#def" target="new">universal, selective, and indicated</a> programming elements. The prevention programs you have in place have a strong basis in the research literature. We encourage you to collect data to evaluate the success of your campus programs and then disseminate your results.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowA1" runat="server" Visible="false">
<!--A- Message<br /><br /> -->
<asp:label id="PA1" runat="server">
Your alcohol prevention programming overall is fairly strong, with a good mix of <a href="calculate.aspx#def" target="new">universal, selective, and indicated</a> programming elements. The prevention programs you are using have an overall strong basis in the research literature. We encourage you to collect data to evaluate the success of your campus programs and then disseminate your results.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowB2" runat="server" Visible="false">
<!--B+ Message<br /><br />-->
<asp:label id="PB2" runat="server">
Your alcohol prevention programming overall is good, with a favorable mix of <a href="calculate.aspx#def" target="new">universal, selective, and indicated</a> programming elements. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score. We encourage you to consult the research and best practices literature to further strengthen your prevention programming.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowB" runat="server" Visible="false">
<!--B Message<br /><br />-->
<asp:label id="PB" runat="server">
Your alcohol prevention programming overall is fairly good. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score. We encourage you to consult the research and best practices literature to further strengthen your prevention programming.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowB1" runat="server" Visible="false">
<!--B- Message<br /><br />-->
<asp:label id="PB1" runat="server">
Your alcohol prevention programming is fair. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score. Additionally, you may have an insufficient amount of programming in some of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We encourage you to consult the research and best practices literature to strengthen your prevention programming.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowC2" runat="server" Visible="false">
<!--C+ Message<br /><br />-->
<asp:label id="PC2" runat="server">
Your alcohol prevention programming is fair, but could be stronger. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score.  Additionally, you may have an insufficient amount of programming in some of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We encourage you to consult the research and best practices literature to strengthen your prevention programming.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowC" runat="server" Visible="false">
<!--C Message<br /><br />-->
<asp:label id="PC" runat="server">
Your alcohol prevention programming is fair, but could be much stronger. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score. Additionally, you may have an insufficient amount of programming in some of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We encourage you to consult the research and best practices literature to strengthen your prevention programming.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowC1" runat="server" Visible="false">
<!--C- Message<br /><br />-->
<asp:label id="PC1" runat="server">
The mix of your alcohol prevention programming is fair, but not yet satisfactory. Some of the prevention programs you are using have a basis in the research literature, but other programs you have in place may lack sufficient evidence, bringing down your overall score.  Additionally, you may have an insufficient amount of programming in some of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We encourage you to consult the research and best practices literature to strengthen your prevention programming.Take heart:  improvement does not necessarily require additional resources.  For instance, it may mean not continuing some of your current programs.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowD2" runat="server" Visible="false">
<!--D+ Message<br /><br />-->
<asp:label id="PD2" runat="server">
The mix of your alcohol prevention programming is not yet satisfactory. Some of the prevention programs you are using may have some research to support their adoption, but others lack an evidence base, bringing down your overall score. Additionally, you may have an insufficient amount of programming in some of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We encourage you to consult the research and best practices literature to strengthen your prevention programming. Take heart:  improvement does not necessarily require additional resources.  For instance, it may mean not continuing some of your current programs.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowD" runat="server" Visible="false">
<!--D Message<br /><br />-->
<asp:label id="PD" runat="server">
The mix of your alcohol prevention programming is not yet satisfactory. Some of the prevention programs you are using may have some research to support their adoption, but many others lack an evidence base, bringing down your overall score. Additionally, you may have an insufficient amount of programming in some or all of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We strongly encourage you to consult the research and best practices literature to strengthen your prevention programming. Take heart:  improvement does not necessarily require additional resources.  For instance, it may mean not continuing some of your current programs.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowD1" runat="server" Visible="false">
<!--D- Message<br /><br />-->
<asp:label id="PD1" runat="server">
The mix of your alcohol prevention programming is not yet satisfactory. Some of the prevention programs you are using may have some research to support their adoption, but many others lack an evidence base, bringing down your overall score. Additionally, you may also have an insufficient amount of programming in some or all of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>) to support healthy student behavior. We strongly encourage you to consult the research and best practices literature to strengthen your prevention programming. Take heart:  improvement does not necessarily require additional resources.  For instance, it may mean not continuing some of your current programs.
<br /><br /></asp:label>
</asp:Panel>
<asp:Panel ID="ShowF" runat="server" Visible="false">
<!--F Message<br /><br />-->
<asp:label id="PF" runat="server">
Your current alcohol prevention programming receives a failing grade. The programs you are using lack a sufficient research base to support their adoption. Additionally, you may also have an insufficient amount of programming in some or all of the prevention categories (<a href="calculate.aspx#def" target="new">universal, selective, indicated</a>). We strongly encourage you to consult the research and best practices literature to strengthen your prevention programming. Take heart:  improvement does not necessarily require additional resources.  For instance, it may mean not continuing some of your current programs.
<br /><br /></asp:label>
</asp:Panel>
</div>
<asp:Label ID="noteHead" runat="server" Visible="false" class="notehead">Additional feedback about your programming:<br /></asp:Label>
<asp:Label ID="notes" runat="server" Visible="false"></asp:Label>
<asp:Label ID="warnings" runat="server" Visible="false"></asp:Label>
</asp:Panel>


<asp:Panel ID="InstScore" runat="server" Visible="false">
 <div class="greenBar">Institutionalization  Grade:</div>
<div id="scoreCard2">
<div class="grade"><asp:label ID="theScore" runat="server"></asp:label></div>
<asp:Label ID="budget" runat="server" Visible="false"></asp:Label>
<asp:Label ID="instDebug" runat="server" Visible="false"></asp:Label>
<asp:Label ID="notes2" runat="server" Visible="false"></asp:Label>
<asp:Label ID="warnings2" runat="server" Visible="false"></asp:Label>
<asp:Panel ID="ShowInstA2" runat="server" Visible="false">
<!--A+ Message<br /><br />-->
<asp:Label ID="IA2" runat="server">
Your institution demonstrates a superior level of commitment to alcohol prevention at an organizational level.<br /><br />
</asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstA" runat="server" Visible="false">
<!--A Message<br /><br />-->
<asp:Label ID="IA" runat="server">
Your institution demonstrates an excellent level of commitment to alcohol prevention at an organizational level.<br /><br />
</asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstA1" runat="server" Visible="false">
<!--A- Message<br /><br />-->
<asp:Label ID="IA1" runat="server">
Your institution demonstrates a very strong level of commitment to alcohol prevention at an organizational level.<br /><br />
</asp:Label></asp:Panel>
<asp:Panel ID="ShowInstB2" runat="server" Visible="false">
<!--B+ Message<br /><br />-->
<asp:Label ID="IB2" runat="server">
Your institution demonstrates a strong level of commitment to alcohol prevention at an organizational level, though there are some additional steps that could be taken to improve your score.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstB" runat="server" Visible="false">
<!--B Message<br /><br />-->
<asp:Label ID="IB" runat="server">
Your institution demonstrates a good level of commitment to alcohol prevention at an organizational level, though there are additional steps that could be taken to improve your score.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstB1" runat="server" Visible="false">
<!--B- Message<br /><br />-->
<asp:Label ID="IB1" runat="server">
Your institution demonstrates a moderate level of commitment to alcohol prevention at an organizational level. Improvement in your score could be achieved with a commitment of more resources or more visible senior leadership on this issue.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstC2" runat="server" Visible="false">
<!--C+ Message<br /><br />-->
<asp:Label ID="IC2" runat="server">
Your institution demonstrates a fair level of commitment to alcohol prevention at an organizational level. Improvement in your score could be achieved with a commitment of more resources or more visible senior leadership on this issue.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstC" runat="server" Visible="false">
<!--C Message<br /><br />-->
<asp:Label ID="IC" runat="server">
Your institution demonstrates a fair level of commitment to alcohol prevention at an organizational level. Improvement in your score could be achieved with a commitment of more resources or more visible senior leadership on this issue.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstC1" runat="server" Visible="false">
<!--C- Message<br /><br />-->
<asp:Label ID="IC1" runat="server">
Your institution demonstrates a minimal level of commitment to alcohol prevention at an organizational level. Improvement in your score could be achieved with a commitment of more resources or more visible senior leadership on this issue.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstD2" runat="server" Visible="false">
<!--D+ Message<br /><br />-->
<asp:Label ID="ID2" runat="server">
Your institution demonstrates a minimal level of commitment to alcohol prevention at an organizational level. Increased resources devoted to alcohol prevention, more visible senior leadership, and stronger collaboration across multiple constituencies are key areas for improvement.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstD" runat="server" Visible="false">
<!--D Message<br /><br />-->
<asp:Label ID="ID0" runat="server">
Your institution demonstrates a weak level of commitment to alcohol prevention at an organizational level. Increased resources devoted to alcohol prevention, more visible senior leadership, and stronger collaboration across multiple constituencies are key areas for improvement.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstD1" runat="server" Visible="false">
<!--D- Message<br /><br />-->
<asp:Label ID="ID1" runat="server">
Your institution demonstrates a very weak level of commitment to alcohol prevention at an organizational level. Increased resources devoted to alcohol prevention, more visible senior leadership, and stronger collaboration across multiple constituencies are key areas for improvement.
<br /><br /></asp:Label>
</asp:Panel>
<asp:Panel ID="ShowInstF" runat="server" Visible="false">
<!--F Message<br /><br />-->
<asp:Label ID="IF0" runat="server">
Your institution is failing to demonstrate a commitment to alcohol prevention at the organizational level. Increased resources devoted to alcohol prevention, more visible senior leadership, and stronger collaboration across multiple constituencies are key areas for improvement.
<br /><br />
</asp:Label>
</asp:Panel>
</div>
    </asp:Panel>

<asp:linkButton ID="otherSurvey" runat="server" class="highlight">How was your grade calculated?</asp:linkButton>
<p></p>
<div class="teal">Recommended Resources: </div>
<p><asp:HyperLink ID="research1" target="_new" runat="server" Visible="true" NavigateUrl="http://www.outsidetheclassroom.com/community/research/alcohol-free-options.aspx" Text="Using Alcohol-Free Options to Promote a Healthy Campus Environment"></asp:HyperLink>
</p>
<p><asp:HyperLink ID="research2" target="_new" runat="server" Visible="true" NavigateUrl="http://www.outsidetheclassroom.com/community/research/institutionalizing-alcohol-prevention.aspx" Text="Institutionalizing Alcohol Prevention: Strategies for Engaging Key Stakeholders and Aligning Objectives"></asp:HyperLink>
</p>
<iframe src="http://www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.campusdiagnostic.com&amp;layout=standard&amp;show_faces=true&amp;width=450&amp;action=recommend&amp;colorscheme=light&amp;height=80" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:450px; height:80px;" allowTransparency="true"></iframe>

<div id="extra"><br />&nbsp;<br /></div>
<asp:Literal ID="openPrinterPage" runat="server" Visible="false" EnableViewState="false"><script language="javascript">                                                                                             openWindow("printerFriendly.aspx");</script></asp:Literal>
</div>
</asp:Content>
