Imports System.Data
Imports AlcoholEdu
Imports AlcoholEdu.Validation
Imports System.Web.UI.HTMLControls.HtmlGenericControl


Public Class login
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents errorDisplay As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents PlaceHolder1 As System.Web.UI.WebControls.PlaceHolder

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Dim iUserAccountId As Integer
    Dim bAnonymousFlag As Boolean
    Dim varNewUserFlag As Boolean
    Dim varLastSegmentId As Integer
    Dim varMediaPlayerId As Integer
    Dim varConnectionSpeedId As Integer
    Dim varErrorCode As Integer
    Dim iClientId As Integer
    Dim varTimeRemaining As Integer
    Dim varLoginTries As Integer
    Dim varTimeOutAssignedCount As Integer
    Dim varLoginType As Integer
    Dim txtUserLogin As String
    Dim txtUserPassword As String
    Dim txtClientLogin As String
    Dim iCourseID As Integer = 4 'Default to HS
    Dim iClientCourseID As Integer
    Dim strMessage As String
    Dim bPortfolio As Boolean = False
    Dim bFollowUp As Boolean = False
    Dim bSchoolEntered As Boolean = False
	Dim bSpecialLogin As Boolean = False
	Dim bStudentTag As Boolean = False
	Dim strStudentTag As String
    Dim lngSegmentID As Integer
    Dim lngSessionID As Long
    Dim encLngSessionID As String
    Dim bNewUser As Boolean
	Dim lc_page_type As String = "login.aspx"
	Dim iStudentLengthID As Integer
    Dim bDemoMode As Boolean = False
    Dim bCustomLogo As Boolean = False
    Dim iSurveyRetakeStatus As Integer = 0
    Dim strClientName As String
    Dim bCustomRegFields As Boolean
    Dim iTiers As Integer


    'TODO: Find a better place to store these error messages.
    Const errmsgINVALID_CLIENT_LOGIN = "You have entered an <b>invalid Login ID</b>.  Please check the materials provided by your administrator and try again or contact your administrator for further assistance."
    Const errmsgINVALID_USERLOGIN = "Please verify that you are entering the correct email address and password.  <b>If you are a ""New User"",</b> please enter your <b>Login ID</b> under ""New User"" to create an account."
    'Const errmsgINVALID_PASSWORD = "The Password you entered is invalid.<br><li>If you are a returning user, please enter your User Id and Password again.  <br><li>If you are a new user, please enter your School Id."
    Const errmsgINVALID_PASSWORD_E1 = "The password you entered does not match the one we have on record for <<UserLogin>>. Please verify that you are entering the correct email address and password and try again."
    Const errmsgUSER_ACCOUNT_EXPIRED = "Your user account has expired.  Please contact your administrator for further assistance."
    Const errmsgINVALID_ANON_PASSWORD = "The password you entered is invalid.  Please enter your email address and password and try again."
    Const errmsgCLIENT_COURSE_DISABLED = "You might be trying to access an inactive course.  Please check the materials provided by your administrator and try again or contact your administrator for further assistance."
    Const errmsgCLIENT_ACCOUNT_DISABLED = "This account has been disabled.  Please check the materials provided by your administrator and try again or contact your administrator for further assistance."
    Const errmsgLOGIN_NOT_UNIQUE = "There is a problem with your Login ID.  Please check the materials provided by your administrator and try again or contact your administrator for further assistance."
    Const errmsgMISSING_INFO = "You did not complete all the information.  Please reenter your information and try again or contact your administrator for further assistance."
    'Do we need the messages below?
    Const errmsgFOLLOW_UP_NOT_READY = "You may not access your account at this time. <b>You will be notified via email</b> when it is time for you to return and complete the rest of the course."
    Const errmsgCOMPLETED_COURSE = "You have completed this course. <b>Your account has been deactivated.</b> If you would like to take the course again, please contact us as admin@outsidetheclassroom.com."
    Const errmsgSAA = "<b>Your school requires that you log in through your school’s web portal.</b>  Please log in to that portal to locate the AlcoholEdu link.  Please contact your administrator if you require further assistance."


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim errMsg As String
        Dim errCode As Integer



        'Before we do anything else, let's set the pageheader propery for this page
        'Dim myHeaderControl = Page.FindControl("pageHeader")
        'myHeaderControl.lc_page_type = lc_page_type
        'Dim myFooterControl = Page.FindControl("pageFooter")
        'myFooterControl.lc_page_type = lc_page_type

        'Check to see if the form has been submitted
        If Page.IsPostBack Or Request.Form("FormMode") = "SUBMIT" Then
            'Load variables passed from user
            txtUserLogin = CleanChar(Trim(Request.Form("txtUserLogin")))
            txtUserPassword = CleanChar(Trim(Request.Form("txtUserPassword")))
            txtClientLogin = CleanChar(Trim(Request.Form("txtClientLogin")))

            'Check for blank form
            If (txtUserLogin = "" Or txtUserPassword = "") And txtClientLogin = "" Then
                errCode = 1
            ElseIf Trim(txtUserLogin) <> "" And Trim(txtUserPassword) <> "" Then             'Returning User
                If Trim(txtClientLogin) <> "" Then
                    bSchoolEntered = True
                End If
                If isValidUserLogin() Then
                    updateSession(iUserAccountId, 0, 0, 0, varLastSegmentId, "", iClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID, bDemoMode, iSurveyRetakeStatus)
                    'TODO: Add bDemoMode to the session at this point
                    enterCourse()
                Else
                    updateSession(iUserAccountId, 0, 0, 0, varLastSegmentId, "", iClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID, bDemoMode, iSurveyRetakeStatus)
                    displayError(bSchoolEntered)
                End If
            ElseIf Trim(txtClientLogin <> "") Then           'New User
                bSchoolEntered = True
                bNewUser = True
                If isValidSchoolLogin() Then
                    updateSession(0, 0, 0, 0, varLastSegmentId, "", iClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID, bDemoMode, iSurveyRetakeStatus)
                    enterCourse()
                Else
                    updateSession(0, 0, 0, 0, varLastSegmentId, "", iClientId, 0, 0, True, "start", iCourseID, 0, 0, iClientCourseID, bDemoMode, iSurveyRetakeStatus)
                    displayError(bSchoolEntered)
                End If
            End If
        Else          'Just display blank page
            'Clear the existing session
            Session.Clear()
            'Add all items to session, but no values yet
            setSession(0, 0, 0, 0, 0, "", iClientId, 0, 0, True, "start", iCourseID, 0, iClientCourseID, iClientCourseID)
            setyear()
        End If

    End Sub
    Private Sub setYear()
        Dim thisYear As Literal
        thisYear = FindControl("thisYear")
        thisYear.Text = DatePart(DateInterval.Year, Today).ToString
    End Sub
    Function isValidUserLogin() As Boolean
        Dim bValidUser As Boolean
        'Run the Validation class to see if this is a valid user login.
        Dim objIsValidUser As New Validation.User.UserValidation
        bValidUser = objIsValidUser.ValidateReturningAEUser(txtUserLogin, txtUserPassword, iCourseID, iUserAccountId, bAnonymousFlag, varNewUserFlag, varLastSegmentId, varMediaPlayerId, varConnectionSpeedId, varErrorCode, iClientId, varTimeRemaining, varLoginTries, varTimeOutAssignedCount, varLoginType, iClientCourseID, bDemoMode, iSurveyRetakeStatus, bCustomLogo)

        If bValidUser = False Then
            Select Case varErrorCode
                Case 6
                    bValidUser = True
                    bPortfolio = True
                    bFollowUp = False
                Case 7
                    bValidUser = True
                    bFollowUp = True
                Case 11
                    bFollowUp = False
                    bPortfolio = True
            End Select
        Else
            If iClientCourseID = 11 Then          'Parent Course, send them error message
                bValidUser = False
                bFollowUp = False
                bPortfolio = False
            End If
        End If
        isValidUserLogin = bValidUser
    End Function

    'Public Function isValidSchoolLogin() As Boolean
    '    Dim bValidLoginID As Boolean
    '    Dim objIsValidUser As New Validation.User.UserValidation
    '    bValidLoginID = objIsValidUser.ValidateAELoginID(iCourseID, txtClientLogin, iClientId, bAnonymousFlag, varErrorCode, bSpecialLogin, iClientCourseID, iStudentLengthID, bDemoMode, strStudentTag, iSurveyRetakeStatus, 1)

    '    If iClientCourseID = 11 Then       'Parent Course
    '        bValidLoginID = False
    '        varErrorCode = 1
    '    End If

    '    isValidSchoolLogin = bValidLoginID
    'End Function

    Public Function isValidSchoolLogin() As Boolean
        Dim bValidLoginID As Boolean
        Dim objIsValidUser As New Validation.User.UserValidation
        Dim bCustomLogo As Boolean = False
        bValidLoginID = objIsValidUser.ValidateAELoginID(iCourseID, txtClientLogin, iClientId, bAnonymousFlag, varErrorCode, bSpecialLogin, iClientCourseID, iStudentLengthID, bDemoMode, strStudentTag, iSurveyRetakeStatus, 1, strClientName, bCustomRegFields, bCustomLogo, iTiers)

        If iClientCourseID = 21 Then       'Parent Course
            bValidLoginID = False
            varErrorCode = 5
        ElseIf iClientCourseID = 7 Then
            bValidLoginID = False
            varErrorCode = 4
        End If
        Session.Item("bCustomLogo") = bCustomLogo

        Return bValidLoginID
    End Function

    Public Sub displayError(ByVal bSchoolEntered)
        If bNewUser Then
            Select Case varErrorCode
                Case 1
                    strMessage = errmsgINVALID_CLIENT_LOGIN
                Case 2
                    strMessage = errmsgCLIENT_COURSE_DISABLED
                Case 3
                    strMessage = errmsgCLIENT_ACCOUNT_DISABLED
                Case 4
                    strMessage = "The LoginID entered is for the other version of AlcoholEdu for High School. Please login here <a href=""http://hs.alcoholedu.com/login.aspx"">http://hs.alcoholedu.com/login.aspx</a>."
                Case 5
                    strMessage = "The LoginID entered is for the Parents version of AlcoholEdu for High Schoo."
            End Select
        Else
            Select Case varErrorCode
                Case 1
                    strMessage = errmsgLOGIN_NOT_UNIQUE
                Case 2
                    strMessage = errmsgINVALID_USERLOGIN
                Case 3
                    strMessage = Replace(errmsgINVALID_PASSWORD_E1, "<<UserLogin>>", txtUserLogin) 'errmsgINVALID_PASSWORD
                Case 4
                    strMessage = errmsgUSER_ACCOUNT_EXPIRED
                    'bPortfolio = true
                    'blnFollowUp = false
                Case 5
                    strMessage = errmsgINVALID_ANON_PASSWORD
                Case 6 'This shouldn't ever display redirect to Notebook
                    'strMessage = errmsgFOLLOW_UP_NOT_READY
                    Call gotoNotebook(varErrorCode)
                Case 8
                    strMessage = errmsgMISSING_INFO
                Case 11 'This shouldn't ever display redirect to Notebook
                    'strMessage = errmsgCOMPLETED_COURSE
                    Call gotoNotebook(varErrorCode)
                Case 13
                    strMessage = errmsgSAA
                Case Else
                    strMessage = errmsgINVALID_USERLOGIN
            End Select
        End If
        If strMessage <> "" Then
            strMessage = "<b><img src=""/images/x.gif"" height=""35"" width=""32"" alt=""ERROR"" border=""0"" hspace=""4"">ERROR!</b><br>" & strMessage
        End If
        errorDisplay.InnerHtml = strMessage
    End Sub

    Public Sub enterCourse()
        Dim strDestinationPage
        Dim bParent As Boolean = False
        If bNewUser Then
            If bAnonymousFlag Then
                strDestinationPage = "NewUserAnonymous.aspx?lngSegmentId=" & lngSegmentID & "&blnSchool=" & bSchoolEntered & "&SIDLen=" & iStudentLengthID
            Else
                Dim strSSLPath As String
                strSSLPath = System.Configuration.ConfigurationManager.AppSettings.Item("SSLPath")
                Session.Item("bTiers") = iTiers
                'For testing locally only. Use a different domain for New User to test SSL and sessions.
                strDestinationPage = "/NewUser.aspx?lngSegmentId=" & lngSegmentID & "&blnSchool=" & bSchoolEntered & "&course=" & iCourseID & "&exID=" & bSpecialLogin & "&Client=" & iClientId & "&blnUniqueUser=true&login=" & txtUserLogin & "&pwd=" & txtUserPassword & "&intlogintype=" & varLoginType & "&SIDLen=" & iStudentLengthID & "&STAG=" & Server.UrlEncode(strStudentTag) & "&pf=" & bParent & "&name=" & Server.UrlEncode("strClientName") & "&bCF=" & Server.UrlEncode(bCustomRegFields.ToString) & "&tier=" & iTiers.ToString()
            End If

        Else
            If bPortfolio Then
                strDestinationPage = "/portal/myalcoholedu.aspx?lngSegmentId=" & lngSegmentID & "&errcode=" & varErrorCode
            ElseIf bSchoolEntered Then
                strDestinationPage = "/message.aspx?SchoolEntered=1&lngSegmentId=" & lngSegmentID
            Else
                strDestinationPage = "/message.aspx?lngSegmentId=" & lngSegmentID
            End If
        End If

        Response.Redirect(strDestinationPage)
    End Sub

    Public Sub setSession(ByVal iUserAccountID As Integer, ByVal iMediaPlayerID As Integer, ByVal iConnectionSpeedID As Integer, ByVal iAssessmentID As Integer, ByVal iSegmentID As Integer, ByVal strSegmentName As String, ByVal iClientID As Integer, ByVal iTestID As Integer, ByVal iScreenHeight As Integer, ByVal bJava As Boolean, ByVal strNavFunction As String, ByVal iCourseID As Integer, ByVal iClassificationID As Integer, ByVal iCourseTypeID As Integer, ByVal iClientCourseID As Integer)
        Dim objSession = New StateMgt.Session
        objSession.setSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
    End Sub

    Public Sub updateSession(ByVal iUserAccountID As Integer, ByVal iMediaPlayerID As Integer, ByVal iConnectionSpeedID As Integer, ByVal iAssessmentID As Integer, ByVal iSegmentID As Integer, ByVal strSegmentName As String, ByVal iClientID As Integer, ByVal iTestID As Integer, ByVal iScreenHeight As Integer, ByVal bJava As Boolean, ByVal strNavFunction As String, ByVal iCourseID As Integer, ByVal iClassificationID As Integer, ByVal iCourseTypeID As Integer, ByVal iClientCourseID As Integer, ByVal bDemoMode As Boolean, ByVal iSurveyRetakeStatus As Integer)
        Dim objSession = New StateMgt.Session
        objSession.updateSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID, bDemoMode, iSurveyRetakeStatus)
    End Sub

    Public Sub readSession(ByRef iUserAccountID As Integer, ByRef iMediaPlayerID As Integer, ByRef iConnectionSpeedID As Integer, ByRef iAssessmentID As Integer, ByRef iSegmentID As Integer, ByRef strSegmentName As String, ByRef iClientID As Integer, ByRef iTestID As Integer, ByRef iScreenHeight As Integer, ByRef bJava As Boolean, ByRef strNavFunction As String, ByRef iCourseID As Integer, ByRef iClassificationID As Integer, ByRef iCourseTypeID As Integer, ByVal iClientCourseID As Integer)
        Dim objSession = New StateMgt.Session
        objSession.readSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
    End Sub

    Public Sub gotoNotebook(ByVal varErrorCode As String)
        Dim strURL As String
        'If iCourseID = 5 Then		  'Sanctions
        'strURL = "/portal/sanctions.aspx?errcode=" & varErrorCode
        'Else
        strURL = "/portal/myalcoholedu.aspx?errcode=" & varErrorCode
        'End If

        Response.Redirect(strURL)
    End Sub

    Public Function CleanChar(ByVal strString As String) As String
        Dim strOutput As String
        strOutput = Replace(strString, "<", "")
        strOutput = Replace(strOutput, ">", "")
        'strOutput = Replace(strOutput, "'", "''")
        Return strOutput

    End Function

End Class
