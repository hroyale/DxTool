
Imports System.Data
Imports AlcoholEdu.Validation
Imports AlcoholEdu.Navigation
Imports AlcoholEdu
Imports System.Web.UI.Control
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.ListBox
Imports System.Web.Security
Imports System.DateTime




Public Class NewUser
    Inherits System.Web.UI.Page

    'initialize variables
    Dim g_strServerDateTime As String
    Dim strUniquePassword As String
    Dim strUniqueUserId As String
    Dim intLoginType As Integer
    Dim bShowSpecial As Object
    Dim strMessage As String
    Dim blnPageOkay As Boolean
    Dim strBirthDate As String
    Dim intCrossClient As Integer
    Dim intCrossClientErr As Integer
    Dim bClone As Boolean
    Dim intSaveErr As Integer
    Dim blnIsLowAcct As Boolean
    Dim blnSentEmail As Boolean
    Dim strDestinationPage As String
    Dim encLngSessionID As String
    Dim iCourseID As Integer
    Dim g_strTimeRemaining As String
    Dim lngOrigUserAccountID As Integer
	Dim lc_page_type As String = "newuser.aspx"
	Dim page_type As String = "newuser.aspx"
    Dim shouldBeLoggedIn As Boolean = False
    Dim strMiddleInitial As String
    Dim intPasswordHintId As Object
    Dim strFirstName As String
    Dim strUserLogin As String
    Dim strPasswordHintAnswer As String
    Dim strPassword As String
    Dim strLastName As String
    Dim intErrorCode As Object
    Dim dtBirthDate As String
    Dim intYearOfStudyId As Object
    Dim bitNewUserFlag As Object
	Dim strGender As String
	Dim iGenderID As Integer
    Dim intUserId As Object
    Dim flagErrorCode As Object
    Dim strOther As String
	Dim strExternalID As String
	Dim strStudentTag As String
    Dim UserAccountID As Integer
    Dim SegmentID As Integer
    Dim strDupeEmailMsg As String
    Dim strOtherErrorMsg As String
    Dim blnSchool As Boolean
    Dim intOrigUserAccountID As Integer = 0
    Dim iClientCourseID As Integer
	Dim strURL As String
	Dim dtToday As DateTime = Now()
	Dim txClientLogin As String
	Dim iStudentIDLength As Integer
	Dim bStudentTag As Boolean = False
    Dim bPreLoad As Boolean = False
    Dim iNewUserAccountID As Integer
	Dim bParent As Boolean
    Dim bSkipTiers As Boolean = False
    Dim strClient As String


    Const flagTRUE As String = "1"
    Const flagFALSE As String = "0"
    Const NO_USER_ACCTS_AVAIL = 1
    Const RAPID_LIMIT = 50
    Const INDEFINITE_BLOCK_SEVERITY = 5
    '---- System List Constants ----
    Const listYEAROFSTUDY = "4"
    Const listPASSHINT = "5"
    Const SYSLIST_BYLIST = "1"
    '---- Unique user query option ----
    Const UNIQUSER_EMAIL = "1"
    Const UNIQUSER_USER = "2"
    Protected WithEvents FirstNameValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents LastNameValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents YOBValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents MOBValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents DOBValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents yearofstudy As System.Web.UI.WebControls.RadioButtonList
	Protected WithEvents YearStudyOtherValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents gender As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents GenderValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents EmailValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents ReemailValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents PasswordValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RepasswordValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents yearstudyother As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents PasswordHintValidator As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents PasswordHintAnswerValidator As System.Web.UI.WebControls.RequiredFieldValidator

    Protected WithEvents radYearOfStudyId As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents GenderChoice As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents txtFirstName As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtMiddleInitial As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtLastName As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtBDateMonth As System.Web.UI.HtmlControls.HtmlSelect
    'Protected WithEvents txtBDateDay As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtBDateYear As System.Web.UI.HtmlControls.HtmlSelect
    'Protected WithEvents CrossClient As System.Web.UI.HtmlControls.HtmlInputCheckBox
	Protected WithEvents iamparent As System.Web.UI.HtmlControls.HtmlInputCheckBox
	Protected WithEvents txtUserLogin As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtReUserLogin As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtPassword As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtRePassword As System.Web.UI.HtmlControls.HtmlInputText
    'Protected WithEvents txtPasswordHintAnswer As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtexternalID As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents txtStudentTag As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents DupeEmailMsg As System.Web.UI.HtmlControls.HtmlGenericControl
	Protected WithEvents Redirect As System.Web.UI.WebControls.Button
    'Protected WithEvents cmdSubmit As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents OtherErrorMsg As System.Web.UI.HtmlControls.HtmlGenericControl
	Protected WithEvents radYearOfStudyId1 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId2 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId3 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId4 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId5 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId6 As System.Web.UI.WebControls.RadioButton
	Protected WithEvents radYearOfStudyId7 As System.Web.UI.WebControls.RadioButton
    Protected WithEvents ddStudentTag As DropDownList
    Protected WithEvents sidDir As Label
    Protected WithEvents StudentTagDir As Label
    Protected WithEvents emaildir As Label
    Protected WithEvents tierDir As Label
    Protected WithEvents tierLabel As Label
    Protected WithEvents email1 As Label

    Dim blnExID As Boolean
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

	End Sub
    Protected WithEvents Check1 As System.Web.UI.WebControls.CheckBox
    Protected WithEvents Submit1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents fname As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents minitial As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents lname As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents bdatemonth As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents bdateday As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents bdateyear As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents yearstudy1 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy2 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy3 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy4 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy5 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy6 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents yearstudy7 As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents male As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents female As System.Web.UI.HtmlControls.HtmlInputRadioButton
    Protected WithEvents MySpan As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents email As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents reemail As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents password As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents repassword As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents hintanswer As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents externalID As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents create As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents CrossClientHTML As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents PasswordHints As System.Web.UI.WebControls.ListBox
    Protected WithEvents UniqueUserYes As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents UniqueUserNo As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents CheckClient As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents StudentIDRequired As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents passwordhint As System.Web.UI.WebControls.ListBox

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		Trace.Write("Load the page")
		Trace.Write(Response.Buffer.ToString)

		Dim lngSessionID As Integer
		Dim intClientId As Integer

		Dim intSegmentID As Integer
		Dim strStudentIDLength As String
		'Dim strStudentTag As String
		Dim strStudentTagLabel As String
        Dim strDupeEmail2 As String = ""
        Dim bPageSubmitted, bEmail As Boolean
        Dim bCustomRegistration As Boolean
        Dim strSIDDir As String
        Dim strTagDir As String
        Dim strEmailDir As String
        Dim strTagDD, strTierLabel, strTierDir As String

		If Request.Form("FormMode") = "SUBMIT" Or Page.IsPostBack Then
			bPageSubmitted = True
		End If

		If Not bPageSubmitted Then
			'Clear the session so we don't re-use the client id.
            'Session.Clear()
            Session.Remove("iClientID")
		End If

		iCourseID = Session.Item("iCourseID")
		If IsNothing(iCourseID) Or iCourseID = 0 Then
			iCourseID = CInt(Request.QueryString("course"))
			Session.Item("iCourseID") = iCourseID
		End If

		intClientId = Session.Item("iClientId")
		If IsNothing(intClientId) Or intClientId = 0 Then
			intClientId = CInt(Request.QueryString("client"))
			Session.Item("iClientId") = intClientId
		End If

		intSegmentID = Session.Item("iSegmentID")
		If IsNothing(intSegmentID) Or intSegmentID = 0 Then
			intSegmentID = CInt(Request.QueryString("lngSegmentID"))
			Session.Item("iSegmentID") = intSegmentID
		End If


        If Not IsNothing(Session.Item("ClientName")) Then
            strClient = Session.Item("ClientName")
        Else
            strClient = Server.UrlDecode(Request.QueryString("name"))
            Session.Item("ClientName") = strClient
        End If

        If Not IsNothing(Session.Item("bCustomReg")) Then
            bCustomRegistration = Session.Item("bCustomReg")
        Else
            bCustomRegistration = CBool(Server.UrlDecode(Request.QueryString("bCF")))
            Session.Item("bCustomReg") = bCustomRegistration
        End If
		'iClientCourseID = Session("iClientCourseID")

		'set the pageheader properyy for this page
        Dim myHeaderControl As secureHeader
        myHeaderControl = Page.FindControl("pageHeader")
        myHeaderControl.lc_page_type = lc_page_type
        Dim myFooterControl As secureFooter
        myFooterControl = Page.FindControl("pageFooter")
        myFooterControl.lc_page_type = lc_page_type

        If IsNothing(iCourseID) Or IsDBNull(iCourseID) Or iCourseID = 0 Then
            iCourseID = 4 'HS
            'iCourseID = Application("courseALCOHOL_EDU")
        End If


		If IsNothing(intClientId) Then
			intClientId = Request.Form.Item("ClientID")
		End If
		If IsNothing(intClientId) Or intClientId = 0 Then Response.Redirect("/login.aspx")

		strUniquePassword = ""
		strUniqueUserId = ""
		intLoginType = Request.QueryString.Item("intLoginType")
		txClientLogin = Request.QueryString.Item("login")
		strStudentIDLength = Request.QueryString.Item("SIDLen")
		If strStudentIDLength = "" Then
			strStudentIDLength = Request.Form("SIDLen")
		End If
		strStudentTagLabel = Request.QueryString.Item("STAG")

		If Trim(strStudentIDLength) <> "" And IsNumeric(strStudentIDLength) Then
			iStudentIDLength = CInt(strStudentIDLength)
		Else
			iStudentIDLength = 0
		End If

		If Trim(strStudentTagLabel) <> "" Then
			strStudentTagLabel = Replace(strStudentTagLabel, "%20", " ")
			bStudentTag = True
		Else
			strStudentTagLabel = Request.Form("STLabel")
			If strStudentTagLabel <> "" Then
				bStudentTag = True
			End If
		End If


		bShowSpecial = Request.QueryString.Item("exID")
		If IsNothing(bShowSpecial) Then
			bShowSpecial = Request.Form("bShowSpecial")
			If Not IsNothing(bShowSpecial) Then
				bShowSpecial = CBool(Request.Form("bShowSpecial"))
			End If
        End If




        If LCase(Request.QueryString.Item("pf")) = "true" Or Request.QueryString.Item("pf") = "1" Then
            bParent = True
            Session.Item("bParent") = True
        ElseIf Session.Item("bParent") Then
            bParent = True
        End If


        If bCustomRegistration Then
            getRegistrationFields(intClientId, strSIDDir, strTagDir, strEmailDir, strTagDD, strTierLabel, strTierDir, bEmail)
        End If
        '*****************************************************************************    
        'THIS BLOCK OF LOGIC CONTROLS HTML OUTPUT TO PAGE WITH <SPAN> TAGS
        'Don't show this option for Sanctions     
        'If iCourseID = 5 Or iCourseID = "5" Then
        '	Dim objCrossClientHTML As HtmlGenericControl
        '	objCrossClientHTML = FindControl("CrossClientHTML")
        '	objCrossClientHTML.Visible = False
        'Else
        'Dim objCrossClientHTML As HtmlGenericControl
        'objCrossClientHTML = FindControl("CrossClientHTML")
        'objCrossClientHTML.Visible = False
        'End If


        'if student ID is required by school
        Dim objStudentIDRequired As HtmlGenericControl
        objStudentIDRequired = FindControl("StudentIDRequired")
        If bShowSpecial = True Or bShowSpecial = "True" Then
            objStudentIDRequired.Visible = True
            Dim objStudentIDLength As HtmlInputHidden
            objStudentIDLength = FindControl("SIDLen")
            objStudentIDLength.Value = strStudentIDLength

            If strSIDDir <> "" Then
                'show the custom dir
                sidDir.Text = strSIDDir
            End If
        Else
            objStudentIDRequired.Visible = False
        End If

        If bStudentTag Then
            Dim objStudentTagFields As HtmlGenericControl
            objStudentTagFields = FindControl("StudentTag")
            objStudentTagFields.Visible = True
            Dim objSTHidden As HtmlInputHidden
            objSTHidden = FindControl("STLabel")
            objSTHidden.Value = strStudentTagLabel
            Dim objSTLabel As Label
            objSTLabel = FindControl("StudentTagLabel")
            objSTLabel.Text = strStudentTagLabel
            If strTagDD <> "" Then
                'Call the sub that will be showing the dd
                makeDD(strTagDD)
            End If

            If strTagDir <> "" Then
                'show the tag dir
                StudentTagDir.Text = strTagDir
            End If
        End If


        If bEmail Then
            'change the email label
            email1.Text = "Email address"

        End If
        If strEmailDir <> "" Then
            'show the tag dir
            emaildir.Text = strEmailDir
        ElseIf bEmail Then
            emaildir.Text = "Please use your email address as your username."
        End If
        'Dim showParent As Panel
        'showParent = FindControl("showParent")
        'If bParent Then
        '	showParent.Visible = True
        '          'CrossClientHTML.Visible = False
        'Else
        '	showParent.Visible = False
        'End If

        '***********************************************************************************

        'showListValues()
        setSchoolValue()

        If Not Page.IsPostBack Then

            Dim iTiers As Integer
            If Not IsNothing(Session.Item("bTiers")) Then
                iTiers = Session.Item("bTiers")
            End If
            If iTiers > 0 Then
                PopulateTierList(intClientId, iTiers)
                bSkipTiers = False
            Else
                bSkipTiers = True
            End If

            If Not bSkipTiers Then
                tierLabel.Text = strTierLabel
                tierDir.Text = strTierDir
            End If

            If Request.Item("blnUniqueUser") = "true" Then
                strUniqueUserId = cleanQS(Request.Item("login"))
                strUniquePassword = Request.Item("pwd")
            End If
            Session.Item("skipState") = bSkipTiers
        End If


        If Not IsNothing(Session.Item("skipTiers")) Then
            bSkipTiers = Session.Item("skipTiers")
        End If

        'if the form has been submitted...
        If Page.IsPostBack Or Request.Form("FormMode") = "SUBMIT" Then

            Dim strParent As String
            strMessage = ""
            'Load variables passed from user
            strUniqueUserId = Request.Form.Item("strUniqueUser")
            strUniquePassword = Request.Form.Item("strUniquePwd")
            strBirthDate = getDOB(Request.Form("txtBDateMonth"), Request.Form("txtBDateYear"))
            intLoginType = CInt(Request.Form.Item("intLoginType"))
            'If bParent Then
            '    strParent = Request.Form.Item("iamparent")
            'End If
            intCrossClient = 0
            strUserLogin = (Request.Form.Item("txtUserLogin"))
            strFirstName = Request.Form.Item("txtFirstName")
            strMiddleInitial = (Request.Form.Item("txtMiddleInitial"))
            strLastName = Request.Form.Item("txtLastName")
            strPassword = Request.Form.Item("txtPassword")
            intPasswordHintId = 0
            strPasswordHintAnswer = ""
            intYearOfStudyId = 0
            strOther = ""
            strGender = 1
            strExternalID = Request.Form.Item("txtexternalID")
            strStudentTag = Request.Form.Item("txtStudentTag")
            If Not IsNothing(Request.Item("blnSchool")) Then
                blnSchool = CBool(Request.Item("blnSchool"))
            End If
            blnPageOkay = pageIsValid(strMessage, bEmail)

            If IsNothing(intLoginType) Or intLoginType = -1 Then             'isuserunique is false
                intLoginType = -1
            End If
            Dim iTierID As Integer = 0
            If Not bSkipTiers Then
                If Not IsNothing(Request.Form("tierList")) Then
                    intClientId = Request.Form("tierList")
                    Session.Item("iClientID") = intClientId
                End If
            End If

            If blnPageOkay Then
                'Has user already completed AEdu this academic year?
                'If intCrossClient = 1 Then				   'user checked box that they completed AEdu in academic year
                '                Response.Redirect("accounttransfer.aspx")
                '            End If

                'Test 3 (worked)

                If IsNumeric(strGender) Then
                    iGenderID = CInt(strGender)
                Else
                    iGenderID = 1
                End If


                If bClone <> True Then
                    intSaveErr = SaveUserAccount()
                End If

                'Test 4 (worked)

                If intSaveErr = NO_USER_ACCTS_AVAIL Then
                    updateSession(UserAccountID, 0, 0, 0, SegmentID, "", intClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID)
                    Response.Redirect("LogoutMessages.aspx?strMessage=msg_E12")
                ElseIf intSaveErr = 2 Or intSaveErr = 3 Then
                    If intSaveErr = 3 Then
                        strDupeEmailMsg = "<b>A user has already created a record with the Student ID that you entered. If you already created an account with us, then you are a Returning User and should enter your email and password at the login screen.  Are you a returning user? "
                        strDupeEmail2 = "<br><br>If you have not created an AlcoholEdu account, please try again and be sure that you are typing in your correct Student ID. Please refer to the instructions provided by your institution or contact your AlcoholEdu Administrator if you do not know your Student ID."
                    Else
                        strDupeEmailMsg = "<b>Your email address matches an email address that was recorded earlier for this school.&nbsp;&nbsp;If you are a returning user you will be taken to the login page.<br> Are you a returning user?"
                    End If
                    strDupeEmailMsg = strDupeEmailMsg & "&nbsp; <INPUT type='button' value='Yes' name='cmdYes' onclick='redirect();' onKeyPress='redirect();'>" & strDupeEmail2 _
                     & "<br><br>If you are trying to retake the course, use the button above to login as a returning user and follow the instructions in the Notebook.</b><br><br>"

                    Dim objDupeEmail As HtmlGenericControl
                    objDupeEmail = FindControl("DupeEmailMsg")
                    objDupeEmail.InnerHtml = strDupeEmailMsg
                ElseIf intSaveErr = 4 Then
                    strOtherErrorMsg = "<b>You entered a Student ID that does not exist in our records. Please try again. If you are certain that you are entering in the correct Student ID, please refer to the instructions provided by your institution or contact your AlcoholEdu Administrator. </b><br><br>"
                    Dim objOtherError As HtmlGenericControl
                    objOtherError = FindControl("OtherErrorMsg")
                    objOtherError.InnerHtml = strOtherErrorMsg
                ElseIf intSaveErr <> 0 Then
                    strOtherErrorMsg = "<b>There was an error saving your information.<br>Please try again or contact your Adminstrator or use our On-line Help.</b><br><br>"
                    Dim objOtherError As HtmlGenericControl
                    objOtherError = FindControl("OtherErrorMsg")
                    objOtherError.InnerHtml = strOtherErrorMsg
                ElseIf strUniqueUserId <> "" And strUniqueUserId <> "" Then
                    'delete the AnonymousUserAccount
                    deleteUniqueUserId(strUniqueUserId, strUniquePassword)
                End If
            Else             'page not Ok
                Dim objOtherError As HtmlGenericControl
                objOtherError = FindControl("OtherErrorMsg")
                objOtherError.InnerHtml = strMessage
            End If


            'Test 5 (worked)


            If (intSaveErr = 0 And blnPageOkay = True) And bSkipTiers = False Then
                '---- HRH
                

                Dim strUserName As String
                strUserName = strFirstName & " " & strLastName

                Session.Item("iuseraccountID") = UserAccountID
                Session.Item("iClientID") = intClientId
                Session.Item("UserName") = strUserName
                Session.Item("bPreLoad") = bPreLoad
                Session.Item("iTierID") = iTierID

                strUserName = strFirstName & " " & strLastName

                Dim strNonSSLPath As String
                strNonSSLPath = System.Configuration.ConfigurationManager.AppSettings.Item("siteURL")

                strURL = "/newuserlogin.aspx?eefg=" & encryptThis(CStr(UserAccountID)) & "&bPreLoad=" & bPreLoad & "&Name=" & Server.UrlEncode(strUserName) & "&client=" & encryptThis(intClientId.ToString)
                Response.Redirect(strURL)

            ElseIf intSaveErr = 0 And blnPageOkay = True Then
                'Progress to the NewUserLogin like on the old page
                Dim strUserName As String
                strUserName = strFirstName & " " & strLastName

                Dim strNonSSLPath As String
                strNonSSLPath = System.Configuration.ConfigurationManager.AppSettings.Item("siteURL")

                strURL = "/newuserlogin.aspx?eefg=" & encryptThis(CStr(UserAccountID)) & "&bPreLoad=" & bPreLoad & "&Name=" & Server.UrlEncode(strUserName) & "&client=" & encryptThis(intClientId.ToString)
                Response.Redirect(strURL)
                'This is an attempt to not bother the browser with the security warning.
                'Response.Clear()
                'Response.Write("<meta http-equiv=""refresh"" content=""0;url=" & strURL & """>l&nbsp;o&nbsp;a&nbsp;d&nbsp;i&nbsp;n&nbsp;g&nbsp;...")
                'Response.End()
            End If
            'Session.Add("bPreLoad", bPreLoad)
            'updateSession(UserAccountID, 0, 0, 0, SegmentID, "", intClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID)
            'Response.Redirect(strDestinationPage)
        Else
            Dim objOtherError As HtmlGenericControl
            objOtherError = FindControl("OtherErrorMsg")
            objOtherError.InnerHtml = strMessage
        End If


    End Sub

    Public Function isUserUnique(ByRef intQryOption As Integer) As Boolean
        Dim blnIsUserUnique As Boolean

        'Run the Validation class to see if this is a valid user login.
        Dim objIsUserUnique As New AlcoholEdu.Validation.User.UserValidation
        blnIsUserUnique = objIsUserUnique.ValidateUniqueUser(intQryOption, strUserLogin, strFirstName, strMiddleInitial, strLastName, strBirthDate, Session("iClientID"), Session("iCourseID"))

        If blnIsUserUnique = True Then
            isUserUnique = True
        Else
            isUserUnique = False
        End If

    End Function

    Function pageIsValid(ByRef strMessage As String, ByVal bEmail As Boolean) As Boolean

        Dim intSeverityLevel As Integer
        Dim intTimeoutMinutes As Integer
        Dim blnValidBirthDate As Boolean
        Dim bPasswordEntered As Boolean = True
        Dim bRePasswordEntered As Boolean = True


        pageIsValid = True
        blnValidBirthDate = True

        If isUserUnique(UNIQUSER_EMAIL) = False Then
            strMessage = "Your email address matches an email address that was recorded earlier for this school.   <br>If you are a returning user you will be taken to the login page. Are you a returning user?"
            strMessage = strMessage & " &nbsp <INPUT type='button' value='Yes' name='cmdYes' onclick='redirect(" & iCourseID & ");' onKeyPress='redirect(" & iCourseID & ");'>"
            pageIsValid = False
        End If

        If pageIsValid = True Then
            strMessage = ""
            If isValueDefined("txtFirstName") = False Then
                strMessage = strMessage & "<li>First Name requires a value."
                pageIsValid = False
            ElseIf isStringValid("txtFirstName") = False Then
                strMessage = strMessage & "<li>First Name is not valid: only allows letters, the apostrophe, the hyphen, the period and the comma."
                pageIsValid = False
            End If
            If isMIValid("txtMiddleInitial") = False Then
                strMessage = strMessage & "<li>Middle initial is not valid: only allows letters."
                pageIsValid = False
            End If
            If isValueDefined("txtLastName") = False Then
                strMessage = strMessage & "<li>Last Name requires a value."
                pageIsValid = False
            ElseIf isStringValid("txtLastName") = False Then
                strMessage = strMessage & "<li>Last Name is not valid: only allows letters, the apostrophe, the hyphen, the period and the comma."
                pageIsValid = False
            End If
            If Not isValueDefined("txtBDateMonth") Or Not isValueDefined("txtBDateYear") Then
                strMessage = strMessage & "<li>Birth Date requires a value."
                pageIsValid = False
            Else
                If isValidDate((Request.Form("txtBDateMonth")), "1", (Request.Form("txtBDateYear"))) = False Then
                    strMessage = strMessage & "<li>Birth Date must be a valid date."
                    blnValidBirthDate = False
                    pageIsValid = False
                End If

                If blnValidBirthDate = True Then
                    If CInt(Trim(Request.Form("txtBDateYear"))) >= CInt(dtToday.Year()) Or CInt(Trim(Request.Form("txtBDateYear"))) < 1900 Then
                        strMessage = strMessage & "<li>Your birth year must be between 1900 and the previous calendar year."
                        pageIsValid = False
                    End If
                End If
            End If
            'If isValueDefined("radYearOfStudyId") = False Then
            '	strMessage = strMessage & "<li>Year of Study requires a valid selection."
            '	pageIsValid = False
            'ElseIf isOtherSelected(Request.Form("radYearOfStudyId")) And Not isValueDefined("yearstudyother") Then
            '	strMessage = strMessage & "<li>Year of Study: Other requires a value when 'Other' is chosen in Year of Study."
            '	pageIsValid = False

            'ElseIf isOtherSelected(Request.Form("radYearOfStudyId")) = False And isValueDefined("yearstudyother") Then
            '	strMessage = strMessage & "<li>Year of Study: You have entered text in the 'Other' field but have also chosen a specific Year of Study. Please make only one choice for the Year of Study."
            '	pageIsValid = False
            'End If
            'If isValueDefined("GenderChoice") = False Then
            '	strMessage = strMessage & "<li>Must select Male or Female."
            '	pageIsValid = False
            'End If

            If isValueDefined("txtUserLogin") = False Then
                If bEmail Then
                    strMessage = strMessage & "<li>'Email Address' requires a value."
                Else
                    strMessage = strMessage & "<li>'Username' requires a value."
                End If
                pageIsValid = False
            End If

            If bEmail = True And isValidEmail(strUserLogin) = False Then
                strMessage = strMessage & "<li>Invalid email address format in 'Your Email Address'."
                pageIsValid = False
            End If
            'If isValueDefined("txtReUserLogin") = False Then
            '	strMessage = strMessage & "<li>'Re-enter Email Address' requires a value."
            '	pageIsValid = False
            'ElseIf isValueMatch("txtUserLogin", "txtReUserLogin") = False Then
            '	strMessage = strMessage & "<li>Your Email Address entry must match the Re-enter Email Address entry."
            '	pageIsValid = False
            'End If
            If isValueDefined("txtPassword") = False Then
                strMessage = strMessage & "<li>Password requires a value."
                pageIsValid = False
                bPasswordEntered = False
            ElseIf isPasswordValid("txtPassword") = False Then
                strMessage = strMessage & "<li>Password is not valid: only allows letters, numbers, the hyphen (-), the period (.) and the underscore (_)."
                pageIsValid = False
            End If
            If isValueDefined("txtRePassword") = False Then
                strMessage = strMessage & "<li>Re-enter Password requires a value."
                pageIsValid = False
                bRePasswordEntered = False
            ElseIf isPasswordValid("txtRePassword") = False Then
                strMessage = strMessage & "<li>Re-enter Password is not valid: only allows letters, numbers, the hyphen (-), the period (.) and the underscore (_)."
                pageIsValid = False
            End If

            If bPasswordEntered And bRePasswordEntered Then
                If isValueMatch("txtPassword", "txtRePassword") = False Then
                    strMessage = strMessage & "<li>Password entry must match the Re-enter Password value."
                    pageIsValid = False
                End If
            End If
            If bPasswordEntered Then
                If isValueLongEnough("txtPassword") = False Then
                    strMessage = strMessage & "<li>Your password must be at least 5 characters long."
                    pageIsValid = False
                End If
            End If
            'If isItemSelected("cboPasswordHintId") = False Then
            'strMessage = strMessage & "<li>Password Hint Question requires a valid selection."
            'pageIsValid = False
            'End If
            'If isValueDefined("txtPasswordHintAnswer") = False Then
            '	strMessage = strMessage & "<li>Password Hint Answer requires a value."
            '	pageIsValid = False
            'End If
            'if isUserUnique(UNIQUSER_USER) = false then
            '	strMessage = strMessage & "<li>User is not unique for this school."
            '	pageIsValid = false
            'end if 
            If bShowSpecial = "True" Or bShowSpecial = True Then
                If Request.Form("txtexternalID") = "" Or Request.Form("txtexternalID") = "0" Then
                    strMessage = strMessage & "<li>Student ID is required."
                    pageIsValid = False
                ElseIf Trim(txClientLogin) <> "" And Trim(Request.Form("txtexternalID")) = Trim(txClientLogin) Then
                    strMessage = strMessage & "<li>Student ID is an ID issued to you by your school, not the course Login ID."
                    pageIsValid = False
                ElseIf checkIDLength() = False Then
                    strMessage = strMessage & "<li>You entered a Student ID that is not the correct length. Please refer to the instructions provided by your institution."
                    pageIsValid = False
                End If

            End If
            'If bParent Then
            '	If isValueDefined("iamparent") = False Then
            '		strMessage = "<br><br>This course is just for Parents. If you are a student, please refer to the instructions directed to you (not your parents) for the proper login id.<br><br>" & strMessage
            '		pageIsValid = False
            '	End If
            'End If

            If pageIsValid = False Then
                strMessage = "Please correct the following error(s) that are present in this page:" & strMessage
            End If
        End If

    End Function




	'save user account info to DB
	Function SaveUserAccount() As Integer
		UserAccountID = 0
		Dim intError As Integer
		'Run the Validation class to see if this is a valid user login.
        Dim objSaveUser As New AlcoholEdu.Validation.User.UserValidation
		'Call the procedure with a 0 at the end to tell it not to search by Name, just Email.
		intError = objSaveUser.SaveUser(Session("iClientID"), UserAccountID, strUserLogin, strPassword, strFirstName, strMiddleInitial, strLastName, intPasswordHintId, strPasswordHintAnswer, strBirthDate, intYearOfStudyId, strOther, iGenderID, strExternalID, SegmentID, strStudentTag, bPreLoad)

		SaveUserAccount = intError
	End Function

	'if user indicated that they completed AEdu in current academic year check DB to verify
	Function CrossClientCheck() As Integer
		Dim intAccountError As Integer
		Dim intUseName As Integer = 0

		'Run the Validation class to see if this is a valid user login.
        Dim objAccountCreated As New AlcoholEdu.Validation.User.UserValidation
		'Call the procedure with a 0 at the end to tell it not to search by Name, just Email.
		intAccountError = objAccountCreated.CrossClientCheck(Session("iCourseID"), strUserLogin, strPassword, strFirstName, strLastName, strBirthDate, intUseName, intOrigUserAccountID)

		CrossClientCheck = intAccountError
	End Function

	'if user did already complete the course under a different name, clone useraccount
    Function CrossClientClone(ByRef iNewUserAccountID As Integer) As Integer
        Dim intCloneError As Integer
        'Run the Validation class to see if this is a valid user login.
        Dim objCloneCreated As New AlcoholEdu.Validation.User.UserValidation
        'Call the procedure with a 0 at the end to tell it not to search by Name, just Email.
        intCloneError = objCloneCreated.CreateClone(Session("iClientID"), Session("iCourseID"), intOrigUserAccountID, iNewUserAccountID)

        If intCloneError = 0 Then
            bClone = True
        Else
            bClone = False
        End If

        CrossClientClone = intCloneError

    End Function

    Function cleanQS(ByRef str2clean As String) As String
        Dim strTemp As String
        strTemp = Replace(str2clean, "%20", " ")
        cleanQS = strTemp
    End Function

    'JS 2/14/04 Note:  could not get data binding to work...so I went with tried and true below
    Function getListValues(ByRef intListId As String) As String
        Dim DRPasswordHints As SqlClient.SqlDataReader
        'get the list values
        Dim objPasswordHints As New AlcoholEdu.Content.PopulateLists
        DRPasswordHints = objPasswordHints.getPasswordHints()

        Dim strHtml As String = ""
        Dim strPasswordHint As String
        Dim strValueField As String

        'js 2/17/05 cannot get databinding to work
        'passwordhint.DataSource = DRPasswordHints
        'passwordhint.DataTextField = DRPasswordHints("vchItemName")
        'passwordhint.DataValueField = DRPasswordHints("iSystemItemId")
        'passwordhint.DataBind()

        'will have to resort to the old method of using html string until databinding gets figured out...
        While DRPasswordHints.Read()
            strPasswordHint = DRPasswordHints("vchPasswordHintType")
            strValueField = DRPasswordHints("iPasswordHintTypeID")
            strHtml = strHtml & "<option value=" & strValueField & ">" & strPasswordHint & "</option>" & vbCrLf
        End While

        If Not IsNothing(DRPasswordHints) Then DRPasswordHints.Close()
        DRPasswordHints = Nothing

        getListValues = strHtml

	End Function

	Sub showListValues()
		Dim objListValues As Literal
		objListValues = FindControl("listValues")
		objListValues.Text = getListValues(5)
	End Sub

	Sub setSchoolValue()
		Dim objSchoolValue As HtmlInputHidden
		objSchoolValue = FindControl("blnSchool")
		objSchoolValue.Value = Request.Item("blnSchool")
	End Sub
	Private Sub deleteUniqueUserId(ByVal userid As String, ByVal password As String)
		'Run the Validation class and delete unique user info
        Dim objDeleteUniqueUser As New AlcoholEdu.Validation.User.UserValidation
		'thought that some kind of error detection would be required here if delete went wrong,
		'but in testing the sproc "spdAnonymousUserAccount", no matter what you feed it, including
		'empty strings, it executes successfully
		objDeleteUniqueUser.DeleteUniqueUser(userid, password)

	End Sub

	Public Sub updateSession(ByVal UserAccountID As Integer, ByVal iMediaPlayerID As Integer, ByVal iConnectionSpeedID As Integer, ByVal iAssessmentID As Integer, ByVal iSegmentID As Integer, ByVal strSegmentName As String, ByVal iClientID As Integer, ByVal iTestID As Integer, ByVal iScreenHeight As Integer, ByVal bJava As Boolean, ByVal strNavFunction As String, ByVal iCourseID As Integer, ByVal iClassificationID As Integer, ByVal iCourseTypeID As Integer, ByVal iClientCourseID As Integer)
        Dim objSession = New AlcoholEdu.StateMgt.Session
		objSession.updateSession(UserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
	End Sub



	Function isOtherSelected(ByVal intOptionVal As Integer) As Boolean
		Dim iOtherValue As Integer = 7		  'Replaced with correct value
		If intOptionVal = iOtherValue Then
			Return True			 'Other is selected
		Else
			Return False			 'Other is not selected
		End If
	End Function

	'=====================================================================================

	Function isValueDefined(ByVal strFieldName As String) As Boolean

		If Trim(Request.Form(strFieldName)) = "" Then
			Return False
		Else
			Return True
		End If

	End Function

	'=====================================================================================

	Function isItemSelected(ByVal strFieldName As String) As Boolean
		'TODO: Don't know if this will work or not. Changed 0 to "" because form values are strings
		If Trim(Request.Form(strFieldName)) = "" Then
			Return False
		Else
			Return True
		End If

	End Function
	'=====================================================================================

	Function isMIValid(ByVal strFieldName As String) As Boolean
		' Checks whether the typed string is valid.
		' 1/9/01

		Dim x As Integer		 'for loop
		Dim blnAnswer As Boolean		  'boolean answer
		Dim strInput As String

		blnAnswer = False		  'default to false

		strInput = Trim(Request.Form(strFieldName))		'Get from form

		Dim aChar		  'As String * 1	'holds the letter to be tested.

		If strInput = "" Or strInput = " " Then
			blnAnswer = True
		Else
			For x = 1 To Len(strInput)
				aChar = Mid(strInput, x, 1)
				If (UCase(aChar) >= "A" And UCase(aChar) <= "Z") Then
					blnAnswer = True
				Else
					blnAnswer = False
					Exit For
				End If
			Next
		End If
		isMIValid = blnAnswer

	End Function


	Function isStringValid(ByVal strFieldName As String) As Boolean
		' Checks whether the typed string is valid.
		' 1/9/01

		Dim x As Integer		  'for loop
		Dim blnAnswer As Boolean		  'boolean answer
		Dim strInput As String

		blnAnswer = False		  'default to false

		strInput = Trim(Request.Form(strFieldName))		'Get from form

		Dim aChar		  'As String * 1	'holds the letter to be tested.

		For x = 1 To Len(strInput)
			aChar = Mid(strInput, x, 1)
			If aChar = "-" Or aChar = "." Or aChar = " " Or aChar = "'" Or aChar = "," Or (UCase(aChar) >= "A" And UCase(aChar) <= "Z") Then
				blnAnswer = True
			Else
				blnAnswer = False
				Exit For
			End If
		Next
		Return blnAnswer

	End Function


	Function isPasswordValid(ByVal strFieldName As String) As Boolean
		' Checks whether the typed string is valid.
		' 1/9/01

		Dim x As Integer		  'for loop
		Dim blnAnswer As Boolean		  'boolean answer
		Dim strInput As String

		blnAnswer = False		  'default to false

		strInput = Trim(Request.Form(strFieldName))		'Get from form

		Dim aChar		  'As String * 1	'holds the letter to be tested.

		For x = 1 To Len(strInput)
			aChar = Mid(strInput, x, 1)
			If aChar = "-" Or aChar = "." Or aChar = "'" Or aChar = "_" Or (UCase(aChar) >= "A" And UCase(aChar) <= "Z") Then
				blnAnswer = True
			ElseIf IsNumeric(aChar) Then
				blnAnswer = True
			Else
				blnAnswer = False
				Exit For
			End If
		Next
		Return blnAnswer

	End Function


	Function isValueMatch(ByVal strFieldName1 As String, ByVal strFieldName2 As String) As Boolean

		If Trim(Request.Form(strFieldName1)) = Trim(Request.Form(strFieldName2)) Then
			Return True
		Else
			Return False
		End If

	End Function

	Function isValueLongEnough(ByVal strFieldName)
		If Len(Request.Form(strFieldName)) < 5 Then
			isValueLongEnough = False
		Else
			isValueLongEnough = True
		End If

	End Function

	Private Function checkIDLength() As Boolean
		Dim strID As String
		strID = Request.Form("txtexternalID")

		If strID <> "" Then
			If iStudentIDLength > 0 Then
				If Len(strID) <> iStudentIDLength Then
					Return False
				Else
					Return True
				End If
			Else
				Return True
			End If
		Else
			Return False
		End If

	End Function

	Function isValidEmail(ByVal strEmail As String) As Boolean
        Dim strEmailValidate, ValidFlag, atCount, SpecialFlag, atLoop, atChr, tAry1
        Dim BadFlag As Boolean
		Dim UserName, DomainName

		strEmailValidate = strEmail

		ValidFlag = False
		If (strEmailValidate <> "") And (InStr(1, strEmailValidate, "@") > 0) And (InStr(1, strEmailValidate, ".") > 0) Then
			atCount = 0
			SpecialFlag = False
			For atLoop = 1 To Len(strEmailValidate)
				atChr = Mid(strEmailValidate, atLoop, 1)
				If atChr = "@" Then atCount = atCount + 1
				If (atChr >= Chr(32)) And (atChr <= Chr(44)) Then SpecialFlag = True
				If (atChr = Chr(47)) Or (atChr = Chr(96)) Or (atChr >= Chr(123)) Then SpecialFlag = True
				If (atChr >= Chr(58)) And (atChr <= Chr(63)) Then SpecialFlag = True
				If (atChr >= Chr(91)) And (atChr <= Chr(94)) Then SpecialFlag = True
			Next
			If (atCount = 1) And (SpecialFlag = False) Then
				BadFlag = False
				tAry1 = Split(strEmailValidate, "@")
				UserName = tAry1(0)
				DomainName = tAry1(1)
				If (UserName = "") Or (DomainName = "") Then BadFlag = True
				If Mid(DomainName, 1, 1) = "." Then BadFlag = True
				If Mid(DomainName, Len(DomainName), 1) = "." Then BadFlag = True
				ValidFlag = True
			End If
		End If

		If BadFlag = True Then ValidFlag = False

		isValidEmail = ValidFlag
	End Function


	Function isValidDate(ByVal intMonth As String, ByVal intDay As String, ByVal intYear As String) As Boolean
		Dim blnValid As Boolean
		Dim iMonth As Integer
		Dim iDay As Integer
		Dim iYear As Integer

		If IsNumeric(intMonth) And IsNumeric(intDay) And IsNumeric(intYear) Then
			iMonth = CInt(intMonth)
			iDay = CInt(intDay)
			iYear = CInt(intYear)
		Else
			iMonth = 0
			iDay = 0
			iYear = 0
			Return False
		End If

		Select Case iMonth
			'make sure that month is valid
		Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
				blnValid = True
			Case Else
				blnValid = False
		End Select

		If blnValid = True Then
			'make sure day of month is valid for that month
			Select Case iMonth
				Case 1, 3, 5, 7, 8, 10, 12				   'months with 31 days
					If iDay < 1 Or iDay > 31 Then
						blnValid = False
					Else
						blnValid = True
					End If

				Case 4, 6, 9, 11				   'months with 30 days
					If iDay < 1 Or iDay > 30 Then
						blnValid = False
					Else
						blnValid = True
					End If

				Case 2				 'feb...allowed 29 in case of leap year
					If iDay < 1 Or iDay > 29 Then
						blnValid = False
					Else
						blnValid = True
					End If
			End Select
		End If

		If blnValid = True Then
			If Len(Trim(intYear)) = 4 Then
				blnValid = True
			Else
				blnValid = False
			End If
		End If

		If blnValid = True Then
			blnValid = IsDate(intMonth & "/" & intDay & "/" & intYear)
		End If

		Return blnValid

	End Function

	Private Sub authenticateNewUser(ByVal iUserAccountID As Integer)
		FormsAuthentication.RedirectFromLoginPage(CStr(iUserAccountID), True)
	End Sub

	Private Function encryptThis(ByVal strString As String) As String
        Dim objEncryptIT As New AlcoholEdu.Content.PopulateLists
		Return objEncryptIT.Encrypt_It(strString)
	End Function

    Private Sub PopulateTierList(ByVal intClientid As Integer, ByVal iTiers As Integer)
        Dim dt As DataTable
        Dim objList As New DataAccess.DAUsers
        If iTiers = 2 Then
            dt = objList.getTierTwoByTeacher(intClientid) 'Really master, not teacher
        Else
            dt = objList.getTierOneByTeacher(intClientid) 'Really master, not teacher
        End If

        Dim tierList As DropDownList
        tierList = FindControl("tierList")
        Dim showTier As Panel
        showTier = FindControl("showTier")

        If dt.Rows.Count() > 0 Then
            tierList.DataSource = dt
            tierList.DataTextField = "vchName"
            tierList.DataValueField = "iClientID"

            ' Bind the data to the control.
            tierList.DataBind()

            ' Set the default selected item, if desired.
            tierList.SelectedIndex = 0
            showTier.Visible = True
        Else
            'This is one of those clients without sub-accounts (Demo??) login without the states.
            tierList.Visible = False
            bSkipTiers = True

            showTier.Visible = False
        End If
    End Sub
    Function getDOB(ByVal strMonth, ByVal strYear)
        Dim dtBDate
        If IsDate(strMonth & "/31/" & strYear) Then
            dtBDate = CDate(strMonth & "/31/" & strYear)
        ElseIf IsDate(strMonth & "/30/" & strYear) Then
            dtBDate = CDate(strMonth & "/30/" & strYear)
        ElseIf IsDate(strMonth & "/29/" & strYear) Then
            dtBDate = CDate(strMonth & "/29/" & strYear)
        ElseIf IsDate(strMonth & "/28/" & strYear) Then
            dtBDate = CDate(strMonth & "/28/" & strYear)
        Else
            dtBDate = CDate("01/01/2003")
        End If

        getDOB = dtBDate

    End Function

    Private Sub PopulateChapterList(ByVal intClientid As Integer, ByVal iStateID As Integer)
        Dim dt As DataTable
        Dim objList As New DataAccess.DAUsers
        Dim iRows As Integer
        dt = objList.getGreekLifeChaptersByState(intClientid, iStateID)
        Dim chapterList As DropDownList
        chapterList = FindControl("chapterList")

        If dt.Rows.Count() > 0 Then
            'Add the blank row to the datatable
            AddBlank(dt)
            iRows = dt.Rows.Count
            chapterList.DataSource = dt
            chapterList.DataTextField = "vchName"
            chapterList.DataValueField = "iClientID"

            ' Bind the data to the control.
            chapterList.DataBind()

            ' Set the default selected item, if desired.
            chapterList.SelectedIndex = iRows - 1

        Else
            OtherErrorMsg.InnerHtml = "No chapters located for the state selected. Please go back and try again."
        End If
    End Sub

    Private Sub AddBlank(ByRef dt As DataTable)
        Dim dr As DataRow = dt.NewRow()
        dr("vchName") = "Choose One"
        dr("iClientID") = "0"

        dt.Rows.Add(dr)

    End Sub

    Sub getRegistrationFields(ByVal iClientID As Integer, ByRef strSidDir As String, ByRef strTagDir As String, ByRef strEmailDir As String, ByRef strTagDD As String, ByRef strTierLabel As String, ByRef strTierDir As String, ByRef bEmail As Boolean)
        Dim objRegFld As New DataAccess.DAClient
        Dim dr As SqlClient.SqlDataReader
        dr = objRegFld.getClientRegFieldsHS(iClientID)
        If dr.Read() Then
            If dr.Item("tiSIDDir") > 0 Then
                strSidDir = dr.Item("vchSIDInstr")
            End If
            If dr.Item("tiTagDir") > 0 Then
                strTagDir = dr.Item("vchTagInstr")
            End If
            If dr.Item("tiTagDD") > 0 Then
                strTagDD = dr.Item("txTagDD")
            End If
            If dr.Item("tiEmailDir") Then
                strEmailDir = dr.Item("vchEmailInstr")
            End If
            If dr.Item("tiStudentEmail") > 0 Then
                'Using email address, change label and check for valid email when submitted
                bEmail = True
            Else
                bEmail = False
            End If
            If Session.Item("bTiers") = 2 Then
                strTierLabel = dr.Item("vchTier2Label")
                strTierDir = dr.Item("vchTier2Dir")
            ElseIf Session.Item("bTiers") = 1 Then
                strTierLabel = dr.Item("vchTier1Label")
                strTierDir = dr.Item("vchTier1Dir")
            End If

        End If
        dr.Close()

    End Sub

    Private Sub makeDD(ByVal strDD As String)
        Dim i As Integer = 0
        txtStudentTag.Visible = False
        ddStudentTag.Visible = True
        'Parse strDD and loop through creating a list item for each one.
        Dim arrItems() As String
        If InStr(strDD, ",") Then
            arrItems = Split(strDD, ",")
        ElseIf InStr(strDD, ";") Then
            arrItems = Split(strDD, ";")
        ElseIf InStr(strDD, vbCrLf) Then
            arrItems = Split(strDD, vbCrLf)
        Else
            txtStudentTag.Visible = True
            ddStudentTag.Visible = False
        End If

        If UBound(arrItems) > 0 Then
            'Let's first add the Select Item
            Dim Li As New ListItem
            Li = makeListItem("Choose One")
            ddStudentTag.Items.Add(Li)
            'ddStudentTag.Items(0).Selected = True
            For i = 0 To UBound(arrItems)
                Dim LItem As New ListItem
                LItem = makeListItem(arrItems(i))
                ddStudentTag.Items.Add(LItem)
            Next
        End If
        ddStudentTag.DataBind()
    End Sub

    Private Function makeListItem(ByVal strChoice As String) As ListItem
        Dim LI As New ListItem
        LI.Text = strChoice
        LI.Value = strChoice 'TODO: Do we want to save the text or the index?
        Return LI
    End Function
End Class