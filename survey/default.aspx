<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="default.aspx.vb" Inherits="DxTool._default1" %>
<%@ import Namespace="System.Data.SQLClient" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="OtcData" %>
<%@ Import Namespace="AlcoholEdu"%>
<%@ Register TagPrefix="AE" TagName="Header" src="\code\headerNoNav.ascx"%>
<%@ Register TagPrefix="AE" TagName="Footer" src="\code\pagefooter.ascx"%>
<AE:HEADER id="pageHeader" runat="server"></AE:HEADER>
<form id="theForm" name="theForm" runat="server">
	<script language="javascript1.2" type="text/Javascript">
<!--
hasAnswered = false;

function setAnsweredSomething() {
	hasAnswered = true;
	answeredObject = getFormObject("answeredFlag");
	answeredObject.value = 1;
}

function maxCharaters(field) {
    /*Validate textarea length before submitting the form*/

    var maxChar = 500;
    if (field.value.length > maxChar) {
        diff=field.value.length - maxChar;
        if (diff>1)
            diff = diff + " characters";
        else
            diff = diff + " character";

        alert("This field is limited to " + maxChar + " characters\n" + "Please reduce the text by " + diff);
        field.focus();
        return (false);
    }
}

function checkSurvey() {
	answeredObject = getFormObject("answeredFlag");
	if (answeredObject.value == 1) {
		hasAnswered = true;
	}

	if (hasAnswered) {
		answeredObject = getFormObject("answeredFlag");
		answeredObject.value = 1;
		return true;
	} else {
		if (confirm("Are you sure you want to proceed to the next page of the survey?\n\nYou have not filled in any responses on this page.")) {
			if (confirm("Please remember that this survey is anonymous and that your survey responses are never linked to you in any way.\n\nYour answers to the survey questions are extremely helpful to us in determining the effectiveness of this course.\n\nPlease go back to the survey and check that you have answered all the questions.")) {
				return false;
			} else {
				return true;
			}
		} else {
			return false;
		}
	}
}

function goToPreviousPage() {
	setNav('back');
	var theFormObj = getObject("theForm");
	submitTheForm();
}

function goToNextPage() {
	if ((checkSurvey()) && (checkRequiredAnswers())) {
		setNav('next');
		var theFormObj = getObject("theForm");
		submitTheForm();
	}
}

function validate(field) {
		var valid = "0123456789"
		var ok = "yes";
		var temp;
		for (var i=0; i<field.value.length; i++) {
		temp = "" + field.value.substring(i, i+1);
		if (valid.indexOf(temp) == "-1") ok = "no";
		}
		if (ok == "no") {
		alert("Please use zero or a positive, whole (i.e., no decimals) number only!");
		//field.value = "";
		field.focus();
		field.select();
		  }
		if (ok == "yes"  ){
			if (field.value > 900){
				ok == "no";
				alert("Please enter a more appropriate number. Your response was too high.");
				field.focus();
				field.select();
				}
			}
		}


function validateCalMinMax(field, min, max) {
	var valid = "0123456789?";
	var doc;
	var ok = "yes";
	var temp;
	for (var i=0; i<field.value.length; i++) {
	temp = "" + field.value.substring(i, i+1);
	if (valid.indexOf(temp) < 0) ok = "no";
	}
	if (ok == "no") {
		alert("Please enter a whole number between " + min + " and " + max + ".");
			//field.focus();
			//field.select();
			//have to use setTimeout() because apparently Netscape 7+
			// has a problem with returning focus to a field following OnBlur
			setTimeout('document.theForm.'+field.name+'.focus()',10);
		}

    //JS 4/16/08 -- since jacey wants '?' in the blank Calendar grid inputs
    // then this presents a problem with validating a '?' against min/max validation...
    // will have to think of a way to deal with this...
	//if (ok == "yes"  ){
	//	if ( field.value < min || field.value > max) {
	//		alert("Please enter a whole number between " + min + " and " + max + ".");
	//		//field.focus();
	//		//field.select();
	//		//have to use setTimeout() because apparently Netscape 7+
	//		// has a problem with returning focus to a field following OnBlur
	//		setTimeout('document.theForm.'+field.name+'.focus()',10);
	//		return false;
	//	}
	// }
}

function validateMinMax(field, min, max) {
	var valid = "0123456789";
	var doc;
	var ok = "yes";
	var temp;
	for (var i=0; i<field.value.length; i++) {
	temp = "" + field.value.substring(i, i+1);
	if (valid.indexOf(temp) < 0) ok = "no";
	}
	if (ok == "no") {
		alert("Please enter a whole number between " + min + " and " + max + ".");
			//field.focus();
			//field.select();
			//have to use setTimeout() because apparently Netscape 7+
			// has a problem with returning focus to a field following OnBlur
			setTimeout('document.theForm.'+field.name+'.focus()',10);
		}
   
	if (ok == "yes"  ){
		if ( field.value < min || field.value > max) {
			alert("Please enter a whole number between " + min + " and " + max + ".");
			//field.focus();
			//field.select();
			//have to use setTimeout() because apparently Netscape 7+
			// has a problem with returning focus to a field following OnBlur
			setTimeout('document.theForm.'+field.name+'.focus()',10);
			return false;
		}
	 }
}

//5/8/08 JS:  added the "}else{return true;" to the top part of this function...
// ...this was to address problem on page where "wieght?" is asked and user
// can input pounds or kilograms...without the }else{ code added the user
// will get the 'needed for customization' popup if they don't have a value
// in BOTH boxes...this new code will alleviate that.
function checkRequiredAnswers(){
	var requiredListObject = getFormObject("reqQuestions");
	var requiredListValues = requiredListObject.value;
	var valid = "true";
	if (requiredListValues.length > 0) {
		if(requiredListValues.indexOf(",") != -1){
			var listArray = requiredListValues.split(",");
			for (var i=0; i<listArray.length; i++) {
				var fieldName = listArray[i];
				var reqAnswerObject = getFormObject(fieldName);
				var reqAnswerValue = reqAnswerObject.value;
				if ((reqAnswerValue == ' ') || (reqAnswerValue.length < 1)){
				valid = "false";
				requiredListObject.value = '';
				}else{
		            return true;
					}
				}
			}else{
			var fieldName = requiredListValues;
			var reqAnswerObject = getFormObject(fieldName);
			var reqAnswerValue = reqAnswerObject.value;
			if ((reqAnswerValue == ' ') || (reqAnswerValue.length < 1)){
				valid = "false";
				requiredListObject.value = '';
					}
			}
		}

	if (valid == "false"){
		alert("You have skipped a question we need to customize AlcoholEdu for you.\n\nWhile you may choose not to answer any question that makes you uncomfortable, by not answering this particular question you may not receive the full benefit of the course.\n\nNote that this survey is anonymous and that your survey responses are never linked to you in any way.\n\nPlease consider going back now and answering this question.");
		return false;
	}else{
		return true;
	}
}

function openSurveyInfo(){
	window.open("/SurveyInfo.html","SurveyInfo","height=500,width=750,titlebar,resizable=yes,scrollbars=yes");
}


// -->

	</script>
	<script language="javascript" type="text/Javascript"></script>
	<span id="errorMessage" runat="server"></span><br>
	<span id="surveyQuestions" runat="server"></span><br>
	<span id="surveyButton" runat="server"></span><br>
	<asp:Panel id="genderDisclaimer" runat="server" visible="false" class="errorMsg" enableviewstate="false">
	<br />* We recognize and appreciate that not all individuals identify within these binary constructs. The purpose of this question (and similar questions that will appear throughout the course) is to calculate your Blood Alcohol Content (BAC), which is based on physiological variables specific to your biological sex and not related to your gender identity. 
	</asp:Panel>
</form>
<AE:FOOTER id="pageFooter" runat="server"></AE:FOOTER>
