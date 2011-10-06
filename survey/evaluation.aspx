<%@ Page Explicit="true" Language="vb" CodeBehind="evaluation.aspx.vb" AutoEventWireup="false" Inherits="DxTool.evaluation" MasterPageFile="~/Site1.Master" %>


<asp:Content id="Content1" ContentPlaceHolderID="mainContent" runat="server" EnableViewState=true>
<script language="javaScript">
    function openPrivacy() {
        window.open("/privacy.html", "PrivacyStatement", "height=500,width=800,titlebar,resizable=yes,scrollbars=yes");
    }
    function toggle(div_id) {
        var el = document.getElementById(div_id);
        if (el.style.display == 'none') { el.style.display = 'block'; }
        else { el.style.display = 'none'; }
    }
    function blanket_size(popUpDivVar) {
        if (typeof window.innerWidth != 'undefined') {
            viewportheight = window.innerHeight;
        } else {
            viewportheight = document.documentElement.clientHeight;
        }
        if ((viewportheight > document.body.parentNode.scrollHeight) && (viewportheight > document.body.parentNode.clientHeight)) {
            blanket_height = viewportheight;
        } else {
            if (document.body.parentNode.clientHeight > document.body.parentNode.scrollHeight) {
                blanket_height = document.body.parentNode.clientHeight;
            } else {
                blanket_height = document.body.parentNode.scrollHeight;
            }
        }
        var blanket = document.getElementById('blanket');
        blanket.style.height = blanket_height + 'px';
        var popUpDiv = document.getElementById(popUpDivVar);
        popUpDiv_height = blanket_height / 2 - 150; //150 is half popup's height
        popUpDiv.style.top = popUpDiv_height + 'px';
    }
    function window_pos(popUpDivVar) {
        if (typeof window.innerWidth != 'undefined') {
            viewportwidth = window.innerHeight;
        } else {
            viewportwidth = document.documentElement.clientHeight;
        }
        if ((viewportwidth > document.body.parentNode.scrollWidth) && (viewportwidth > document.body.parentNode.clientWidth)) {
            window_width = viewportwidth;
        } else {
            if (document.body.parentNode.clientWidth > document.body.parentNode.scrollWidth) {
                window_width = document.body.parentNode.clientWidth;
            } else {
                window_width = document.body.parentNode.scrollWidth;
            }
        }
        var popUpDiv = document.getElementById(popUpDivVar);
        window_width = window_width / 2 - 150; //150 is half popup's width
        popUpDiv.style.left = '200px';
        popUpDiv.style.top = '200px';
    }
    function popup(windowname) {
        //blanket_size(windowname);
        window_pos(windowname);
        toggle('blanket');
        toggle(windowname);
    }    
</script>
<style type="text/css"> 

#blanket {
   background-color:#339999;
   opacity: 0.5;
   position:absolute;
   filter: alpha(opacity=50);
   z-index: 9001; /*ooveeerrrr nine thoussaaaannnd*/
   top:0px;
   left:0px;
   width:100%;
   height:1400px;

}
.popupq
{
    position:fixed;
    top:100px;
    left:200px;    
	background-color:#FFFFFF;
	width:350px;
	height:320px;
	padding:10px;
	border:solid 2px black;
	text-align:  left;
	z-index: 9002; /*ooveeerrrr nine thoussaaaannnd*/
}
#popUpDiv a, #popUpDiv2 a, #popUpDiv3 a, #popUpDiv4 a, #popUpDiv5 a, #popUpDiv6 a  
{
font-weight: bold;
padding:10px;
}
#backgrounds 
{
    /*text-align:center;*/
    width:450px;
    margin:0px auto 0px auto;
}
#plan3, #plan1, #plan2 
{
    width:130px;
    float:left;    
    }
   
#showStatement
{
    width:612px;
    padding-top:230px;
    margin: 0px auto 0px auto;
    } 
       
.plan3 
{
    background-image:url("/images/plan3.png");
    background-repeat:no-repeat;
    }
.plan2 
{
    background-image:url("/images/plan2.png");
    background-repeat:no-repeat;
        padding-top:230px;
    }    
.plan1 
{
    background-image:url("/images/plan1.png");
    background-repeat:no-repeat;
    }    
   
#statementContent 
{
    width:375px;
    margin:0px auto 0px auto;
    padding:0px;
    }   
       
#btnNext 
{
    padding: 0px 10px 0px 230px;
    }

</style>
  <div id="blanket" style="display:none;"></div>
    <a href="http://www.outsidetheclassroom.com" target="_blank"><asp:Image ID="surveyImage" runat="server" ImageUrl="/images/Survey1.png" AlternateText="The Campus Diagnostic" /></a>
 <asp:label ID="subTitle" runat="server" class="title" Visible="false"></asp:label> 
<asp:Literal ID="alldivtags" runat="server"></asp:Literal>
	<span id="errorMessage" runat="server" class="error"/><span id="evaluationQuestions" runat="server" />
<br><br>

<asp:Label id="custText" runat="server" visible="false"></asp:Label>
<input type="hidden" id="skipFlag" value="" runat="server"><input type="hidden" runat="server" name="answeredFlag" id="answeredFlag" value="" />

<input class="submitButton" type="submit" value="Submit Responses and Continue" id="button1" name="button1" runat="server" visible="true"><input class="submitButton" type="button" value="Submit Responses and Continue" id="button2" name="button2" runat="server" onClick="goToNextPage();" visible="false">


<script language="javascript" type="text/javascript" >
    var hasAnswered = false;

    function setAnsweredSomething(subqid) {
        hasAnswered = true;
        //alert("in answered something");
        subqid = "h" + subqid;
        if (subqid != "hundefined") {
            //alert(subqid);
            popup(subqid);
        }
    }

    function setHasAnswer() {
        answeredObject = getFormObject("answeredFlag");
        answeredObject.value = 1;
    }


    function checkSE() {
        var answeredObject = getFormObject("answeredFlag");
        var answered = answeredObject.value;
        if (answered != 1) {
            alert("Please indicate that you want to attend and/or plan events or activities before selecting them.");
        }
    }


    function maxCharaters(field) {
        /*Validate textarea length before submitting the form*/

        var maxChar = 500
        if (field.value.length > maxChar) {
            diff = field.value.length - maxChar;
            if (diff > 1)
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
        //JS 6/19/03: changed this from (answeredObject.value == 1) to line below...
        // In Netscape 'answeredObject.value' would equate to 1 if a
        // user was returning to the page with previously marked answers; however,
        // 'hasAnswered' would consistently equate to 'false' causing the alert box to appear.
        // Netscape appears to like this better...?
        if (answeredObject.value > 0) {
            hasAnswered = true;
        }

        if (hasAnswered) {
            answeredObject = getFormObject("answeredFlag");
            answeredObject.value = 1;
            return true;
        } else {
            if (confirm("Are you sure you want to proceed to the next page?\n\nYou have not filled in any responses on this page.")) {
                if (confirm("This survey protects your privacy in every way, and the responses are extremely helpful to us in determining the effectiveness of this course.\n\nWould you please fill in the answers?")) {
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
        submitTheForm();
    }

    function goToNextPage() {
        if (checkSurvey()) {
            submitTheForm();
        }
    }

    function validate(field) {
        var valid = "0123456789"
        var ok = "yes";
        var temp;
        for (var i = 0; i < field.value.length; i++) {
            temp = "" + field.value.substring(i, i + 1);
            if (valid.indexOf(temp) == "-1") ok = "no";
        }
        if (ok == "no") {
            alert("Please use zero or a positive, whole (i.e., no decimals) number only!");
            //field.value = "";
            field.focus();
            field.select();
        }
        if (ok == "yes") {
            if (field.value > 900) {
                ok == "no";
                alert("Please enter a more appropriate number. Your response was too high.");
                field.focus();
                field.select();
            }
        }
    }


    function validateMinMax(field, min, max) {
        var valid = "0123456789";
        var ok = "yes";
        var temp;
        for (var i = 0; i < field.value.length; i++) {
            temp = "" + field.value.substring(i, i + 1);
            if (valid.indexOf(temp) < 0) ok = "no";
        }
        if (ok == "no") {
            alert("Please enter a number between " + min + " and " + max + ".");
            field.focus();
            field.select();
        }

        if (ok == "yes") {
            if (field.value < min || field.value > max) {
                alert("Please enter a number between " + min + " and " + max + ".");
                field.focus();
                field.select();
                return false;
            }
        }
    }

    function checkRequiredAnswers() {
        var requiredListObject = getFormObject("reqQuestions");
        var requiredListValues = requiredListObject.value;
        var valid = "true";
        if (requiredListValues.length > 0) {
            if (requiredListValues.indexOf(",") != -1) {
                var listArray = requiredListValues.split(",");
                for (var i = 0; i < listArray.length; i++) {
                    var fieldName = listArray[i];
                    var reqAnswerObject = getFormObject(fieldName);
                    var reqAnswerValue = reqAnswerObject.value;
                    if ((reqAnswerValue == ' ') || (reqAnswerValue.length < 1)) {
                        valid = "false";
                        requiredListObject.value = '';
                    }
                }
            } else {
                var fieldName = requiredListValues;
                var reqAnswerObject = getFormObject(fieldName);
                var reqAnswerValue = reqAnswerObject.value;
                if ((reqAnswerValue == ' ') || (reqAnswerValue.length < 1)) {
                    valid = "false";
                    requiredListObject.value = '';
                }
            }
        }

        if (valid == "false") {
            alert("You have skipped a question that is required in order to customize AlcoholEdu for you. Consider going back now and answering that question.  Please know that nobody will ever connect your answers to your name, and your individual responses will never be identified.");
            return false;
        } else {
            return true;
        }
    }

    // -->

</script>

</asp:Content>
