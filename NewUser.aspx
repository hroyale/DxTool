<%@ Page explicit="true" CodeBehind="NewUser.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="HighSchoolNew.NewUser" %>
<%@ Register TagPrefix="AE" TagName="Footer" src="code\securefooter.ascx"%>
<%@ Register TagPrefix="AE" TagName="Header" src="code\secureheader.ascx"%>

<AE:HEADER id="pageHeader" runat="server"></AE:HEADER>
<form id="form1" name="form1" runat="server">
	<p><b>Before you begin</b>, please create an account profile.</p>
	

	
	<TABLE cellSpacing="0" cellPadding="2" width="100%" border="0">
		<tr>
			<td colSpan="3"><SPAN id="DupeEmailMsg" runat="server" enableViewState="False"></SPAN>
		<SPAN id="OtherErrorMsg" runat="server" style="error" class="error" enableViewState="False"></SPAN></td>
		</tr>

		<TR>
			<TD class="blue" width="27%"><B><LABEL accessKey="D" for="fname">First Name</LABEL></B></TD>
			<TD><INPUT id="txtFirstName" maxLength="75" runat="server" NAME="txtFirstName"></TD>
			<TD>&nbsp;</TD>
		</TR>

		<TR>
			<TD class="blue"><LABEL accessKey="E" for="minitial">Middle Initial</LABEL></TD>
			<TD><INPUT id="txtMiddleInitial" maxLength="1" size="3" name="txtMiddleInitial" runat="server"></TD>
			<TD>&nbsp;</TD>
		</TR>
		<TR>
			<TD class="blue"><B><LABEL accessKey="F" for="lname">Last Name</LABEL></B></TD>
			<TD><INPUT id="txtLastName" maxLength="75" name="txtLastName" runat="server"></TD>
			<TD>&nbsp;</TD>
		</TR>
		
  <TR>
    <TD style="blue" class="blue"><B>Birth&nbsp;Month&nbsp;and&nbsp;Year&nbsp;</B></TD>
    <TD><LABEL for="bdatemonth" Accesskey="G"><Select name="txtBDateMonth" id="txtBDateMonth" runat="server">
    	<option value="01">January</option>
    	<option value="02">February</option>
    	<option value="03">March</option>    	
    	<option value="04">April</option>
    	<option value="05">May</option>
    	<option value="06">June</option>
    	<option value="07">July</option>
    	<option value="08">August</option>
    	<option value="09">September</option>
    	<option value="10">October</option>    	    	    	    	    	    	    	
    	<option value="11">November</option>
    	<option value="12">December</option>     	
    </Select></LABEL> <!--onBlur="validate(this);"-->
		  <LABEL for="bdateyear" Accesskey="I"><Select id="txtBDateYear" name="txtBDateYear" runat="server">
	    <option value="1998">1998</option>
	    <option value="1997">1997</option>
	    <option value="1996">1996</option>		
	    <option value="1995">1995</option>
	    <option value="1994">1994</option>	
	    <option value="1993">1993</option>	
	    <option value="1992">1992</option>		  
    	<option value="1991">1991</option>		  
    	<option value="1990">1990</option>
	    					
    </Select>		  
		  </LABEL></TD>
		  <TD style="font-size:8pt;" valign=top><b>You must be 13 years or older to take this course.</b></TD>
		  </TR>


			<TR>
				<TD colSpan="3"><br>
				</TD>
			</TR>
			<TR>
				<TD class="blue"><B><asp:label runat="server" id="email1">Username</asp:label>:</B></TD>
				<TD><input id="txtUserLogin" maxLength="150" size="25" name="txtUserLogin" runat="server" onblur="checkForComcast(this.value);"></TD>
				<TD vAlign="top"  class="footer"><b><asp:label id="emailDir" runat="server">Note: Create a username or use the one provided to you.</asp:label>  
						<br><br>
					</b>
				</TD>
			</TR>


		<TR>
			<TD class="blue"><B><LABEL accessKey="P" for="password">Password</LABEL></B></TD>
			<TD><INPUT id="txtPassword" type="password" maxLength="150" name="txtPassword" runat="server">&nbsp;<br><br>
			</TD>
			<TD vAlign="top" rowSpan="2" class="footer"><b>Your password must be at least 5 
					characters long.<br>
				</b>
			</TD>
		</TR>
		<TR>
			<TD class="blue"><B><LABEL accessKey="Q" for="repassword">Re-enter 
      Password</LABEL></B></TD>
			<TD><INPUT id="txtRePassword" type="password" maxLength="150" name="txtRePassword" runat="server">&nbsp;
			</TD>
		</TR>
		<TR>
			<TD colSpan="3">&nbsp;</TD>
		</TR>
		
		<asp:Panel id="showTier" runat="server" visible="false">
        <tr><TD class="blue"><B><asp:Label id="tierLabel" runat="server">Tier 1</asp:Label></B></TD>
            <td>
             <asp:DropDownList id="tierList" runat="server" visible="true"></asp:DropDownList>   
        </td>
        <td vAlign="top" class="footer"><b><asp:Label id="tierDir" runat="server"></asp:Label></b></td>
        </tr>		
		</asp:Panel>
		
		
		<!--if Stuent Tag required show this block -->
		<SPAN id="StudentTag" name="StudentTag" runat="server" visible="false">
			<TR>
				<TD class="blue" vAlign="top"><B><asp:label id="StudentTagLabel" runat="server"></asp:label></B></TD>
				<TD vAlign="top"><input id="txtStudentTag" type="text" maxLength="50" size="20" name="txtStudentTag" runat="server">
				<asp:DropDownList id="ddStudentTag" runat="server" visible="false"></asp:DropDownList>
				</TD>
				<TD vAlign="top" class="footer"><b><asp:label id="StudentTagDir" runat="server">Check your instructions for details about this field.</asp:label></b><br><br></TD>
			</TR>
			<input type="hidden" id="STLabel" name="STLabel" value="1" runat="server">
		</SPAN>
		
		
		<!--if studentID required show this block -->

		<SPAN id="StudentIDRequired" name="StudentIDRequired" runat="server" visible="false">
			<TR>
				<TD class="blue" vAlign="top"><B><LABEL accessKey="T" for="externalID">Student ID</LABEL></B></TD>
				<TD vAlign="top"><input id="txtexternalID" type="text" maxLength="20" size="20" name="txtexternalID" runat="server">
				</TD>
				<TD vAlign="top" class="footer"><b><asp:label id="sidDir" runat="server">Refer to the directions provided by your 
						Administrator for helpful hints on where to find your Student ID.</asp:label></b><br><br><input type="hidden" id="bShowSpecial" name="bShowSpecial" value="true">
						<input type="hidden" id="SIDLen" name="SIDLen" value="0" runat="server"></TD>
			</TR>
						
		</SPAN>
		
						
		
	
<tr><td colspan="3">
<br /><br />
	<asp:button id="continue" runat="server" text="Continue" style="background-color:#4682b4;color:#FFFFFF;" visible="False"/>	
    <asp:ImageButton id="continue2" runat="server" imageurl="/images/CreateMyAccount.jpg" visible="true" hover="/images/CreateMyAccount_A.jpg" class="rollover"/>
</td></tr>

</TABLE>
<div align="right">Note: Fields in <B>BOLD</B> are required.&nbsp; &nbsp;<br />
<br />
<br /></div> 
	<Input type="hidden" id="blnSchool" name="blnSchool" runat="server">
	<Input type="hidden" id="Hidden1" name="blnSchool" runat="server">
	<div id="emailPopUp" runat="server">
		<a style="font-size: 12px; align:right;" onclick="document.getElementById('emailPopUp').style.display = 'none' " onmouseover="this.style.cursor=&quot;pointer&quot; " onfocus="this.blur();"><span style="text-decoration: underline;">X</span></a>
<br />
	It looks like you entered a Comcast email address. Please be sure to add "alcoholedu.com" to your approved senders list or use a different email address.<br /><br />
	<a style="font-size: 12px; align:right;" onclick="document.getElementById('emailPopUp').style.display = 'none' " onmouseover="this.style.cursor=&quot;pointer&quot; " onfocus="this.blur();"><span style="text-decoration: underline;">Close</span></a>
	</div>
<div id="agePopUp" runat="server">
		<a style="font-size: 12px; align:right;" onclick="document.getElementById('agePopUp').style.display = 'none' " onmouseover="this.style.cursor=&quot;pointer&quot; " onfocus="this.blur();"><span style="text-decoration: underline;">X</span></a>
<br />
It looks like you entered an incorrect birth date. Please check the year.
	<a style="font-size: 12px; align:right;" onclick="document.getElementById('agePopUp').style.display = 'none' " onmouseover="this.style.cursor=&quot;pointer&quot; " onfocus="this.blur();"><span style="text-decoration: underline;">Close</span></a>
	</div>
	
	
	<script language="JavaScript">

	    function submitPage() {
	        document.forms['form1'].action = "NewUser.aspx";
	        document.forms['form1'].submit();
	    }

	    function entsub(event) {
	        if (event && event.which == 13) {
	            submitPage();
	        }
	    }

	    function enterSubmit(event) {
	        if (event && event.which == 13) {
	            submitPage();
	        } else if (window.event && window.event.keyCode == 13) {
	            submitPage();
	        }
	    }

	    function checkForComcast(val) {
	        if (val.indexOf("comcast") > 0) {
	            document.getElementById('emailPopUp').style.display = 'block';
	        }
	    }

	    function checkYear(val) {
	        var currentTime = new Date()
	        if ((parseInt(val) + 5) >= currentTime.getFullYear()) {
	            document.getElementById('agePopUp').style.display = 'block';
	        }
	    }

	    function showOptions() {
	        //alert("in show options!");
	        document.getElementById('RegWarning').style.display = 'block';
	    }



	    function redirect(course) {
	        if (course == "5") {
	            document.forms['form1'].action = "http://sanctions.alcoholedu.com/enter.aspx";
	        } else {
	            document.forms['form1'].action = "http://college.alcoholedu.com/login.aspx";
	        }
	        document.forms['form1'].__VIEWSTATE.name = 'NOVIEWSTATE';
	        document.forms['form1'].submit();
	    }

	    function redirectSan() {
	        document.forms['form1'].action = "http://sanctions.alcoholedu.com/enter.aspx";
	        document.forms['form1'].__VIEWSTATE.name = 'NOVIEWSTATE';
	        document.forms['form1'].submit();
	    }

	    function clearText() {
	        document.forms['form1'].elements['txtOther'].value = ""
	    }

	</script>
</form>
<AE:FOOTER id="pageFooter" runat="server"></AE:FOOTER>
