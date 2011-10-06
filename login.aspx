<%@ Page Explicit="true" Language="vb" CodeBehind="login.aspx.vb" AutoEventWireup="False" Inherits="HighSchoolNew.login" validateRequest="false" EnableViewStateMac="false"%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>AlcoholEdu: High School Login</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="stylesheet" href="/documentStyle.css" type="text/css" />
    <link rel="stylesheet" href="styles/loginCLIstyle.css" type="text/css" />

        <SCRIPT type="text/javascript" src="/code/jquery-1.4.2.min.js"></SCRIPT>
         <SCRIPT type="text/javascript" src="/code/main.js"></SCRIPT>	
<script type="text/javascript" src="/code/loginCLI.js">
</script>
</head>

<body>

<div id="stage">

<div id="mainContainer">

<div id="nestedContainer">

<div id="loginTemplate">

<!--top buttons-->



<!--end top buttons-->

<!--bottom left buttons-->

<!--end bottom left buttons-->

<form id="theForm" name="theForm" runat="server">

<input type="hidden" name="prevPage" value="login.aspx" /> 
<input type="hidden" value="1" name="LoginStepNo" /> 
<!-- removed hidden field for blnSchoolEntered -->

<input type="hidden" value="SUBMIT" name="FormMode" /> 


<div id="redLogin">
<div class="loginContainer">

<div id="loginfield">
<b>Login ID</b><br />
<input type="text" style="width: 130px" onkeypress="entsub(event);" id="School" name="txtClientLogin" />
</div>
<a href="javascript:popup('/loginID.html');"><img id="idBtn" src="/images/WhatsanID.gif" alt="What is an ID?" class="rollover" hover="/images/WhatsanID_A.gif" border="0" /></a>

<asp:Imagebutton ID="newUserBtn" runat="server" AlternateText="Log In" ImageUrl="/images/SignUp.gif" class="rollover" hover="/images/SignUp_A.gif" />


</div><!--loginContainer-->
</div><!--redLogin-->

<!--login boxes-->
<div id="blueLogin">

<div class="loginContainer">

<b>Email address</b><br />
<div id="emailfield"><input type="text" style="width: 200px" onblur="tabNext('txtUserPassword');" name="txtUserLogin" id="Email" /></div>

<div id="passfield"><b>Password</b><br /><input type="password" style="width: 120px" onkeypress="entsub(event);" name="txtUserPassword" id="Password" />
</div>

<asp:Imagebutton ID="retUserBtn" runat="server" AlternateText="Log In" ImageUrl="/images/Login.gif" class="rollover" hover="/images/Login_A.gif" visible="true" />

<a class="loginLink" href="lostpassword.html" target="_new"><img src="/images/ForgotPassword.gif" alt="Forgot Password" class="rollover" hover="/images/ForgotPassword_A.gif" border="0" id="lostPwd" /></a>

</div><!--blueContainer-->
</div><!--blueLogin-->



</form>

<div ID="blueNote">
<div id="errorDisplay" runat="server">
<b>Important Note:</b> If you have been asked to complete AlcoholEdu for High School, please refer to the directions you received for your start date; you will not be able to access the course until that date.
</div>

</div><!--blueNote-->

<asp:HyperLink ID="helpBtnLnk" runat="server" NavigateUrl="help.html" Target="_new"><img src="/images/ProblemsLoggingIn.gif" border="0" class="rollover" hover="/images/ProblemsLoggingIn_A.gif" id="helpBtn" /></asp:HyperLink>
<asp:HyperLink ID="reqBtnLnk" runat="server" NavigateUrl="settings.html" Target="_new"><img src="/images/MinimalTechnical.gif" border="0" class="rollover" hover="/images/MinimalTechnical_A.gif" id="reqBtn" /></asp:HyperLink>
<asp:HyperLink ID="settingsBtnLnk" runat="server" NavigateUrl="settings.html" Target="_new"><img src="/images/ReccommendedBrowserSettings.gif" border="0" class="rollover" hover="/images/ReccommendedBrowserSettings_A.gif" id="settingsBtn" /></asp:HyperLink>

</div><!--loginTemplate-->
<div id="copyrightDiv2">Copyright ©2000-<asp:Literal ID="thisYear" runat="server">2010</asp:Literal>, Outside The Classroom, Inc. | <a class="privacy" href="javascript:popup('/privacy.html');">Privacy Statement</a></div>

</div><!-- nested container -->

</div><!--mainContainer-->


</div><!--stage-->



<script language="javascript1.2" type="text/javascript">
<!-- //
	document.forms[0].txtUserLogin.focus();

	function goToLostPassword()
	{
		//var strUserLog;
		//objUserLog = getFormObject("txtUserLogin");
		//strUserLog = objUserLog.value;
		window.location = "LostPassword.aspx"
		//popup("http://supportcenteronline.com/ics/support/default.asp?deptID=713");		
	}

	function entsub(event) {
	 if(event && event.which == 13){
		login();
	} else if (window.event && window.event.keyCode == 13){
		login();
	}

    }

    function tabNext(nextfield){
		eval('document.forms[0].' + nextfield + '.focus()');
    }

	function login(){

		var objScreenHeight;
		var objFormMode;
		objFormMode = getFormObject('FormMode');
		objFormMode.value = 'SUBMIT';
		objScreenHeight = getFormObject('intScreenHeight');
		//objScreenHeight.value = screen.height;
		submitTheForm();
	}

	function popup(p){
		window.open(p,"popupP","height=500,width=800,titlebar,resizable=yes,scrollbars=yes");
	}
// -->

</script>

</body>

</html>
