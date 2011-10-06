

Imports System.Data
Imports AlcoholEdu.DataAccess
Imports AlcoholEdu.Navigation
Imports AlcoholEdu.StateMgt
Imports System.Web.UI.HtmlControls
Partial Public Class _default1
    Inherits System.Web.UI.Page




    'initialize variables    
    Dim iErrorCode As Integer 'Used in reading the XML file
    Dim page_type As String
    Dim intlastPage As Integer
    Dim i As Integer
    Dim pageNumber As Integer
    Dim responseSelected As Boolean
    Dim FreeResponse As String
    Dim answeredSomething As Byte
    Dim subQuestionChoiceID As Integer
    Dim intQuestionFormatID As Integer
    Dim intQuestionTypeID As Integer
    Dim intQuestionNumber As Integer
    Dim icount As Integer
    Dim intQuestionID As Integer
    Dim intAnswerFlag As Integer
    Dim strTextQuestion As String
    Dim strQuestions As String
    Dim intTestTypeID As Integer
    Dim intIdcounter As Integer
    Dim intAccessKey As Integer
    Dim strAccessKey As String
    Dim strReqQuestions As String
    Dim surveyIntro As String
    Dim intSubQuestionID As Integer
    Dim strStyle As String
    Dim bSkipQues As Boolean
    Dim bSkipMoveNext As Boolean
    Dim strQuestionChoice As String
    Dim bPageIsValid As Boolean
    Dim strQuestionText As String
    Dim strMsg As String
    Dim intSurveyType As Integer
    Dim bSurvey As Boolean
    Dim bSpecialFormat As Boolean
    Dim intChoicesCount As Integer
    Dim intNoChoices As Integer
    Dim intMax As Integer
    Dim intSkipPattern As Integer
    Dim bskip As Boolean
    Dim bPremat As Boolean
    Dim bSanctions As Boolean
    Dim bPrematControl As Boolean
    Dim bControl As Boolean
    Dim bDefault As Boolean
    Dim intPageSkipBack As Integer
    Dim intPageBegin As Integer
    Dim intPageSkipForward As Integer
    Dim strHTMLQuestions As String
    Dim strSubQuestions As String
    Dim intFieldCount As Integer
    Dim intDataViewQuestionID As Integer
    Dim intHoldDataViewQuestionID As Integer
    Dim intDataViewIncrement As Integer
    Dim strHtmlButton As String
    Dim intCheckFinal As Integer
    Dim strLinks As String
    Dim navFunction As String
    Dim intPageStart As Integer
    Dim strURL As String
    Dim blnEndSurvey As Boolean
    Dim lc_page_type As String
    Dim intNoFields As Integer
    Dim j As Integer
    Dim intFromPage As Integer
    Dim blnSkipPattern1 As Boolean
    Dim blnSkipPattern2 As Boolean
    Dim blnSkipPattern3 As Boolean
    Dim blnSkipPattern4 As Boolean

    'skip pattern 5 is for the UPenn 2008 control study
    Dim blnSkipPattern5 As Boolean

    'skip patterns 6 - 10 for Qset skips
    Dim blnSkipPattern6 As Boolean 'Drinking Experinces' Qset
    Dim blnSkipPattern7 As Boolean 'Diary' Qset
    Dim blnSkipPattern8 As Boolean 'Diary' Qset
    Dim blnSkipPattern9 As Boolean 'Diary' Qset
    Dim blnSkipPattern10 As Boolean 'Diary' Qset
    Dim blnSkipPattern11 As Boolean 'Alcohol Settings' Qset
    Dim blnSkipPattern12 As Boolean 'Alcohol Settings' Qset
    Dim blnSkipPattern13 As Boolean 'Alcohol Settings' Qset
    Dim blnSkipPattern14 As Boolean 'Energy Drinks' Qset


    '--JS 1/18 for dropdowns
    Dim vchDropDown As String

    'session values
    Dim iUserAccountID As Integer
    Dim iClientCourseID As Integer
    Dim iSegmentID As Integer
    Dim iMediaPlayerID As Integer
    Dim iConnectionSpeedID As Integer
    Dim iAssessmentID As Integer
    Dim strSegmentName As String
    Dim iClientID As Integer
    Dim iTestID As Integer
    Dim iScreenHeight As Integer
    Dim bJava As Boolean
    Dim strNavFunction As String
    Dim iCourseID As Integer
    Dim iClassificationID As Integer
    Dim iCourseTypeID As Integer
    Dim intIOrder As Integer
    Dim blnCustomQuestion As Boolean
    Dim intCustomQuestionNumber As Integer
    Protected WithEvents errorMessage As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents surveyQuestions As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents surveyButton As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents genderDisclaimer As Panel

    Const formatFlag = 1

    'constant values for the ordinal references used in dataset
    Const vchFreeResponse As Integer = 23
    Const iQuestionChoiceID As Integer = 15
    Const vchQuestionChoice As Integer = 16
    Const iFreeResponseID As Integer = 11
    Const iMaximum As Integer = 28
    Const iMinimum As Integer = 27
    Const iResponse As Integer = 22
    Const chSelectionNumber As Integer = 17
    Const iSubQuestionID As Integer = 18
    Const iQColumnWidth As Integer = 13
    Const txQuestion As Integer = 26
    Const iColumnWidth As Integer = 19
    Const iQuestionID As Integer = 0
    Const bComplete As Integer = 9
    Const iQuestionFormatID As Integer = 5


#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

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
        Dim lngSessionID As Integer = Session("SessionID")
        Dim intClientId As Integer = Session("iClientID")
        Dim intCourseID As Integer = Session("iCourseID")
        Dim intSegmentID As Integer = Session("iSegmentID")
        Dim bTimeSurvey As Boolean = True
        Dim startTime As TimeSpan
        Dim answerTable As Hashtable = New Hashtable

        If bTimeSurvey Then
            startTimer(startTime)
        End If

        'JS 4/18/08: ALL HARDCODED QUESTION/CHOICE IDs LINES DENOTED WITH "!!!"

        'TESTING CONTEXT:
        'addItemToContext()

        intTestTypeID = GetTestTypeID(intSegmentID)
        'this is for the pageheader, so appropriate stuff will show
        If intTestTypeID = 11 Then    'followup New TestType
            lc_page_type = "/survey/followup.aspx"
            page_type = "/survey/followup.asp"
        Else    'pre/post
            lc_page_type = "/survey/default.aspx"
            page_type = "/survey/Survey.aspx"
        End If

        'set the pageheader property for this page
        Dim myHeaderControl As headerNoNav
        myHeaderControl = Page.FindControl("pageHeader")
        myHeaderControl.lc_page_type = lc_page_type
        Dim myFooterControl As PageFooter
        myFooterControl = Page.FindControl("pageFooter")
        myFooterControl.lc_page_type = lc_page_type

        intAnswerFlag = 0
        bPageIsValid = True

        'get session data for use on page
        readSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)

        If iTestID <> 6 Then    'Not in the correct Segment Type to be using this page. Send to Error.
            Response.Redirect("/error.aspx?msg=back")
        End If

        If iCourseTypeID = 0 Then
            iCourseTypeID = getCourseTypeID()
        End If
        pageNumber = 0
        strReqQuestions = ""


        If Request.Form("pageNumber") <> "" Then
            pageNumber = CInt(Request.Form("pageNumber"))
        End If

        'check to see if user did not complete survey previously, if so, erase pre-existing answers

        If pageNumber = 0 Then Call checkSurvey() 'Skip forward if done, clear answers if started

        'If this is the first page of the Post Survey, check to make sure the Final Exam has been completed
        If intTestTypeID = 10 And pageNumber = 0 Then    'New TestTypeID
            'intCheckFinal = checkFinal() 'js 6/26/08 no longer do this as with 9.0 postsurvey precedes Final Exm
        Else    'assume this is the PreSurvey
            If pageNumber = 0 Then Call runPrePopulate()
            If pageNumber = 1 And intTestTypeID = 1 Then
                'This will show the disclaimer for the gender question only on the first page of Survey 1.
                genderDisclaimer.Visible = True
            Else
                genderDisclaimer.Visible = False

            End If
        End If


        'if not intro page then save responses
        'Put code here to not save the text only answers too.
        If pageNumber <> 0 And Request.Form("navFunction") <> "back" And Session.Item("skipSave") <> True Then
            Call saveSurvey()
        End If

        'SKIP PATTERNS
        '====================PLP COLLEGE PRE/POST SURVEYS==========================================
        If Request.Form("blnSkipPattern1") <> "" Then
            blnSkipPattern1 = Request.Form("blnSkipPattern1")
        End If
        If Request.Form("blnSkipPattern2") <> "" Then
            blnSkipPattern2 = Request.Form("blnSkipPattern2")
        End If
        If Request.Form("blnSkipPattern3") <> "" Then
            blnSkipPattern3 = Request.Form("blnSkipPattern3")
        End If
        If Request.Form("blnSkipPattern4") <> "" Then
            blnSkipPattern4 = Request.Form("blnSkipPattern4")
        End If
        '=============?????? UPENN SURVEY 3 SKIP PATTERN  ?????  USED ANYMORE ?????===========================================
        If Request.Form("blnSkipPattern5") <> "" Then
            blnSkipPattern5 = Request.Form("blnSkipPattern5")
        End If
        '===================================================================================

        '=================QSET SKIP PATTERNs======================================================
        If Request.Form("blnSkipPattern6") <> "" Then
            blnSkipPattern6 = Request.Form("blnSkipPattern6")
        End If
        If Request.Form("blnSkipPattern7") <> "" Then
            blnSkipPattern7 = Request.Form("blnSkipPattern7")
        End If
        If Request.Form("blnSkipPattern8") <> "" Then
            blnSkipPattern8 = Request.Form("blnSkipPattern8")
        End If
        If Request.Form("blnSkipPattern9") <> "" Then
            blnSkipPattern9 = Request.Form("blnSkipPattern9")
        End If
        If Request.Form("blnSkipPattern10") <> "" Then
            blnSkipPattern10 = Request.Form("blnSkipPattern10")
        End If
        If Request.Form("blnSkipPattern11") <> "" Then
            blnSkipPattern11 = Request.Form("blnSkipPattern11")
        End If
        If Request.Form("blnSkipPattern12") <> "" Then
            blnSkipPattern12 = Request.Form("blnSkipPattern12")
        End If
        If Request.Form("blnSkipPattern13") <> "" Then
            blnSkipPattern13 = Request.Form("blnSkipPattern13")
        End If
        If Request.Form("blnSkipPattern14") <> "" Then
            blnSkipPattern14 = Request.Form("blnSkipPattern14")
        End If
        '====================================================================================

        If Request("navFunction") = "back" Then
            pageNumber = checkSkipBack()
        Else
            If intTestTypeID = 10 Then
                pageNumber = checkSurvey2SkipForward() 'This is just for Survey 2 since we can skip the first question
            ElseIf pageNumber < 2 Then    'don't do checking if not on page 2 
                pageNumber = pageNumber + 1
            Else
                pageNumber = checkSkipForward()
            End If
        End If

        'get Survey questions
        strHTMLQuestions = getSurveyQuestions(pageNumber)

        If Request("EndSurvey") = "1" Then
            blnEndSurvey = True
        End If

        If blnEndSurvey And (Request("navFunction") <> "back") Then
            'need to update strNavFunction in session so that it will redirect correctly
            strNavFunction = "next"
            updateSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
            'Set dtEnd in assessment
            Call setSurveyDate()
            'Remove the datatable with the answers from the session.
            Session.Remove("SurveyAnswers")
            'If the calendar grid stored it's answers, remove those too.
            RemoveCalSession()

            'This is survey 3. Call the procedure that does the classifications for Module 4.
            Dim iTestFormID As Integer
            iTestFormID = getTestFormID()
            Select Case iTestFormID
                Case 42, 113, 917, 918
                    socClassify(iUserAccountID)
                Case 6, 112, 912, 913, 914
                    ClassifyV9Path()
            End Select


            Dim objChangeSegment As Navigation.ChangeSegment = New Navigation.ChangeSegment
            strURL = objChangeSegment.getNextURL()
            If Trim(strURL) <> "" Then
                Response.Redirect(strURL)
            Else
                Response.Redirect("/error.aspx")
            End If
        Else
            'output questions for display    
            Dim objSurveyQuestions As HtmlGenericControl
            objSurveyQuestions = FindControl("surveyQuestions")
            objSurveyQuestions.InnerHtml = strHTMLQuestions

            'output button for display    
            Dim objSurveyButton As HtmlGenericControl
            objSurveyButton = FindControl("surveyButton")
            If pageNumber > 1 Or blnCustomQuestion Then    'if its a custom question page then we don't want the button to read 'Begin Survey'...
                strHtmlButton = "<br><label for=""button1""><input class=""submitButton"" type=""button"" value=""Submit Responses and Continue"" onClick=""goToNextPage()"" id=""button1"" name=""button1""></label>"
            Else
                strHtmlButton = "<br><center><label for=""button1""><input class=""submitButton"" type=""submit"" value=""Begin Survey"" id=""button1"" name=""button1""></label></center>"
            End If
            objSurveyButton.InnerHtml = strHtmlButton
        End If

        If bTimeSurvey Then
            endTimer(startTime)
        End If

    End Sub
    Public Function checkSkipBack()
        Dim pageNumberBack As Integer
        Dim blnSkipBack As Boolean = False
        Dim blnSkip As Boolean
        Dim intPageSkipBackFrom As Integer

        'skip pattern 1:  didn't drink in past year' 
        'Q#5 survey 1/Q#4 survey 3 pre/post; Q#3 survey 1/3 premat control   '...
        If Request.Form("blnSkipPattern1") <> "" Then
            blnSkipPattern1 = Request.Form("blnSkipPattern1")
            If blnSkipPattern1 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                Select Case intTestTypeID
                    Case 1    'survey 1
                        If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Or iCourseTypeID = 4 Then 'premat/saedu premat/summer/saedu summer
                            intPageSkipBackFrom = 8
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 5
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                            'ElseIf iCourseTypeID = 4 Then       'premat control
                            '    intPageSkipBackFrom = 7
                            '    'If pageNumber = intPageSkipBackFrom Then
                            '    '    pageNumber = pageNumber - 4
                            '    '    blnSkipPattern1 = False
                            '    '    blnSkipBack = True
                            '    'End If
                        ElseIf iCourseTypeID = 8 Then       'sanctions
                            intPageSkipBackFrom = 7
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                        Else       'all other courses   'postmat/postmat control
                            intPageSkipBackFrom = 8
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 5
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                        End If
                    Case 10    'survey 2
                        If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                            intPageSkipBackFrom = 3
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            Else
                                intPageSkipBackFrom = 5
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern1 = False
                                    blnSkipBack = True
                                End If
                            End If

                        ElseIf iCourseTypeID = 4 Then       'premat control
                            intPageSkipBackFrom = 0
                        ElseIf iCourseTypeID = 8 Then       'sanctions
                            intPageSkipBackFrom = 7
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                        Else       'all other courses  'postmat
                            intPageSkipBackFrom = 3
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            Else
                                intPageSkipBackFrom = 5
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern1 = False
                                    blnSkipBack = True
                                End If
                            End If
                        End If
                    Case 11    'survey 3
                        If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                            intPageSkipBackFrom = 8
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 4
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                            'ElseIf iCourseTypeID = 4 Then       'premat control
                            '    intPageSkipBackFrom = 7
                            '    If pageNumber = intPageSkipBackFrom Then
                            '        pageNumber = pageNumber - 4
                            '        blnSkipPattern1 = False
                            '        blnSkipBack = True
                            '    End If
                        ElseIf iCourseTypeID = 8 Then       'sanctions
                            intPageSkipBackFrom = 7
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                        Else       'all other courses...postmat/postmat control
                            intPageSkipBackFrom = 8
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 4
                                blnSkipPattern1 = False
                                blnSkipBack = True
                            End If
                        End If
                End Select
            Else
                blnSkipPattern1 = False
                blnSkipBack = False
            End If
        Else
            blnSkipPattern1 = False
        End If

        'skip pattern 2:  for 'no drinks last two weeks' 
        If Not (blnSkipBack) Then
            If Request.Form("blnSkipPattern2") <> "" Then
                blnSkipPattern2 = Request.Form("blnSkipPattern2")
                If blnSkipPattern2 Then 'only proceed if skip pattern = true
                    'set page numbers based on course/testtype
                    Select Case intTestTypeID
                        Case 1    'survey 1
                            If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                                intPageSkipBackFrom = 6
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 4 Then       'premat control
                                intPageSkipBackFrom = 6
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 6
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses   'postmat/postmat control
                                intPageSkipBackFrom = 6
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            End If
                        Case 10    'survey 2
                            If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                                intPageSkipBackFrom = 8
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 4 Then       'premat control
                                intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 10
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses  'postmat
                                intPageSkipBackFrom = 10
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            End If
                        Case 11    'survey 3
                            If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                                intPageSkipBackFrom = 7
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 4 Then       'premat control
                                intPageSkipBackFrom = 6
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 6
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses postmat/postmat control
                                intPageSkipBackFrom = 7
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 2
                                    blnSkipPattern2 = False
                                    blnSkipBack = True
                                End If
                            End If
                    End Select
                Else
                    blnSkipPattern2 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern2 = False
            End If
        End If

        'skip pattern 3:  US citizen?  
        If Not (blnSkipBack) Then
            If Request.Form("blnSkipPattern3") <> "" Then
                blnSkipPattern3 = Request.Form("blnSkipPattern3")
                If blnSkipPattern3 Then 'only proceed if skip pattern = true
                    Select Case intTestTypeID
                        Case 1    'survey 1
                            If iCourseTypeID = 2 Or iCourseTypeID = 16 Then 'premat/summer/
                                intPageSkipBackFrom = 11
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 1
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                                'ElseIf iCourseTypeID = 4 Then       'premat control
                                '    intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 13 Or iCourseTypeID = 12 Or iCourseTypeID = 17 Then 'PreMat/PostMat/Summer SAEDU
                                intPageSkipBackFrom = 12
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 1
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses:  postmat/postmat control
                                intPageSkipBackFrom = 11
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 1
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            End If
                        Case 10    'survey 2
                            If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                                intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 4 Then       'premat control
                                intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 3 Then       'control
                                intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 9
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses
                                intPageSkipBackFrom = 9
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            End If
                        Case 11    'survey 3
                            If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            ElseIf iCourseTypeID = 4 Then       'premat control
                                intPageSkipBackFrom = 0
                            ElseIf iCourseTypeID = 8 Then       'sanctions
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            Else       'all other courses: postmat/ control postmat
                                intPageSkipBackFrom = 13
                                If pageNumber = intPageSkipBackFrom Then
                                    pageNumber = pageNumber - 5
                                    blnSkipPattern3 = False
                                    blnSkipBack = True
                                End If
                            End If
                    End Select
                Else
                    blnSkipPattern3 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern3 = False
            End If
        End If

        'skip pattern 4: user answered 'no' on Q#5 on Survey 1 -- only applies to PostMat Survey 2
        If Request.Form("blnSkipPattern4") <> "" Then
            blnSkipPattern4 = Request.Form("blnSkipPattern4")
            If blnSkipPattern4 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                Select Case intTestTypeID
                    Case 10    'survey 2
                        If iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                            intPageSkipBackFrom = 3
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern4 = False
                                blnSkipBack = True
                            End If
                        ElseIf iCourseTypeID = 4 Then       'premat control
                            intPageSkipBackFrom = 0
                        ElseIf iCourseTypeID = 8 Then       'sanctions
                            intPageSkipBackFrom = 7
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern4 = False
                                blnSkipBack = True
                            End If
                        Else       'all other courses  'postmat
                            intPageSkipBackFrom = 3
                            If pageNumber = intPageSkipBackFrom Then
                                pageNumber = pageNumber - 2
                                blnSkipPattern4 = False
                                blnSkipBack = True
                            End If
                        End If
                End Select
            Else
                blnSkipPattern4 = False
                blnSkipBack = False
            End If
        Else
            blnSkipPattern4 = False
        End If

        '=======================UPENN SURVEY 3 Q#26 SKIP PATTERN=========================
        'skip pattern 5: user input '0' into Q#26 input box 
        If Request.Form("blnSkipPattern5") <> "" Then
            blnSkipPattern5 = Request.Form("blnSkipPattern5")
            If blnSkipPattern5 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 4 Then       'premat control
                    intPageSkipBackFrom = 12
                    If pageNumber = intPageSkipBackFrom Then
                        pageNumber = pageNumber - 3 'go back to pg 9 Q#26
                        blnSkipPattern5 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern5 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern5 = False
            End If
        End If
        '============================END UPENN SURVEY 3 Q#26 SKIP PATTERN=======================


        '=======================QSET SKIP PATTERNs=============================================
        'skip pattern 6: user chose 'no' question #1 Qset 'Drinking Experiences'
        If Request.Form("blnSkipPattern6") <> "" Then
            blnSkipPattern6 = Request.Form("blnSkipPattern6")
            If blnSkipPattern6 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 9
                    If pageNumber = intPageSkipBackFrom Then
                        pageNumber = pageNumber - 7 'go back to pg 2 Q#1
                        blnSkipPattern6 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern6 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern6 = False
            End If
        End If

        'skip pattern 7: user chose 'no' question 1, Qset 'Diary'
        If Request.Form("blnSkipPattern7") <> "" Then
            blnSkipPattern7 = Request.Form("blnSkipPattern7")
            If blnSkipPattern7 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 10 'go back from pg 10 on both survey 1 qset and survey 3 qset; same questionID, same skip pattern
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 8 'go back to pg 2 Q#1 from page 10
                        blnSkipPattern7 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern7 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern7 = False
            End If
        End If

        'skip pattern 8: user chose 'no' question 8, Qset 'Diary'
        If Request.Form("blnSkipPattern8") <> "" Then
            blnSkipPattern8 = Request.Form("blnSkipPattern8")
            If blnSkipPattern8 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 8 'Qset 'Diary question - #23/pg 8 on both survey 1 qset and survey 3 qset; same questionID, same skip pattern
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 4 'go back to pg 4 Q#8 from page 8
                        blnSkipPattern8 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern8 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern8 = False
            End If
        End If

        'skip pattern 9: user chose 'no' question 13, Qset 'Diary'
        If Request.Form("blnSkipPattern9") <> "" Then
            blnSkipPattern9 = Request.Form("blnSkipPattern9")
            If blnSkipPattern9 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 8 'Qset 'Diary' skip patterns on both survey 1 qset and survey 3 qset; same questionID, same skip pattern
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 3 'go back to pg 5 Q#13 from page 8
                        blnSkipPattern9 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern9 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern9 = False
            End If
        End If

        'skip pattern 10: user chose 'no' question 18, Qset 'Diary'
        If Request.Form("blnSkipPattern10") <> "" Then
            blnSkipPattern10 = Request.Form("blnSkipPattern10")
            If blnSkipPattern10 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 8 'Qset 'Diary' skip patterns question - #23/pg 8 on both survey 1 qset and survey 3 qset; same questionID, same skip pattern
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 2 'go back to pg 6 Q#13 from page 8
                        blnSkipPattern10 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern10 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern10 = False
            End If
        End If

        'skip pattern 11: user chose 'no' question 1, Qset 'Alcohol Settings'
        If Request.Form("blnSkipPattern11") <> "" Then
            blnSkipPattern11 = Request.Form("blnSkipPattern11")
            If blnSkipPattern11 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 9 'Qset 'Alcohol Settings' skip patterns question - #1 
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 7 'go back to pg 2
                        blnSkipPattern11 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern11 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern11 = False
            End If
        End If

        'skip pattern 12: user chose 'no' question 5, Qset 'Alcohol Settings'
        If Request.Form("blnSkipPattern12") <> "" Then
            blnSkipPattern12 = Request.Form("blnSkipPattern12")
            If blnSkipPattern12 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 9 'Qset 'Alcohol Settings' skip patterns question - #5 
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 4 'go back to pg 5
                        blnSkipPattern12 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern12 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern12 = False
            End If
        End If

        'skip pattern 13: user chose 'no' question 7, Qset 'Alcohol Settings'
        If Request.Form("blnSkipPattern13") <> "" Then
            blnSkipPattern13 = Request.Form("blnSkipPattern13")
            If blnSkipPattern13 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 9 'Qset 'Alcohol Settings' skip patterns question - #7 
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 3 'go back to pg 7
                        blnSkipPattern13 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern13 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern13 = False
            End If
        End If

        'skip pattern 14: user chose 'no' question 2, Qset 'Energy Drinks'
        If Request.Form("blnSkipPattern14") <> "" Then
            blnSkipPattern14 = Request.Form("blnSkipPattern14")
            If blnSkipPattern14 Then 'only proceed if skip pattern = true
                'set page numbers based on course/testtype
                If iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Then 'premat/saedu premat/summer/saedu summer
                    intPageSkipBackFrom = 5 'Qset 'last page -- 5 -- holds the "thank you for taking this...'
                    If pageNumber = intPageSkipBackFrom Then 'we need to skip back...
                        pageNumber = pageNumber - 3 'go back to pg 2
                        blnSkipPattern14 = False
                        blnSkipBack = True
                    End If
                Else
                    blnSkipPattern14 = False
                    blnSkipBack = False
                End If
            Else
                blnSkipPattern14 = False
            End If
        End If
        '============================END QSET SKIP PATTERNs=======================


        If blnSkipBack = True Then
            checkSkipBack = pageNumber
        Else
            checkSkipBack = pageNumber - 1
        End If
    End Function

    Public Function checkSkipForward()
        Dim strQuestionID As String
        Dim arrQuestionArray() As String
        Dim intQuestionID As Integer
        Dim arrTextAnswerArray() As String
        Dim strTextQuestionID As String
        Dim strQuestion As String
        Dim i As Integer
        Dim theAnswerArray() As String
        Dim skippableArray(14, 14) As Integer
        Dim skippableTextArray(1) As Integer
        Dim x As Integer
        Dim y As Integer
        Dim k As Integer
        Dim t As Integer
        Dim intQuestionChoice As Integer
        Dim blnSkipForward As Boolean = False
        Dim strAnsQuestions As String
        Dim intAnsQuestionsLength As Integer
        Dim strTextAnswer As String
        Dim intTextAnswer1 As Integer
        Dim intTextAnswerTotal1 As Integer
        Dim intTextAnswer2 As Integer
        Dim intTextAnswerTotal2 As Integer
        Dim intCalendarQuestion As Integer
        Dim intSurvey2SkipQuestion As Integer
        '================for upenn control skip============
        Dim intUPennTextQuestion As Integer
        '============================================

        'note: all places where hardcoded questionIDs appear on this page are denoted by:  !!!

        'need to remove last comma as it was causing an empty string to be appended to question list...
        intAnsQuestionsLength = Len(Request.Form.Item("ansQuestions"))
        strAnsQuestions = Left(Request.Form.Item("ansQuestions"), (intAnsQuestionsLength - 1))
        'array for non-text answers
        arrQuestionArray = Split(strAnsQuestions, ",")
        'array for text answers
        arrTextAnswerArray = Split(Request.Form.Item("ansTextQuestions"), ",")
        strTextQuestionID = strQuestionID & "text" & Request.Form.Item(strQuestionID)

        '======='skip pattern 5:  for UPENN CONTROL ONLY...Q#26 Survey 3 (note: there are 2 Q#26; and intro text 'question and the real question which is what we need here==================================
        skippableTextArray(0) = 7335  '8015 on athena:' "During NSO (Thursday thru Monday), on how many days did you drink?"
        intUPennTextQuestion = 7335  '8015 on athena:
        '====================================================================================================================================================

        'create an array to hold skippable question IDs

        'skip pattern 1 -- surveys 1 and 3
        skippableArray(0, 0) = 6211    'Q#5 survey 1(pre/post), Q#3 survey 1(preControl), Q#4 survey 3; during past year, consumed any alcohol? !!!
        skippableArray(0, 1) = 25033  'choice = 'No' !!!  
        'skip pattern 2 -- surveys 1 and 3
        skippableArray(1, 0) = 6229  'Q#8 survey 1(pre/post), Q#4 survey 1(preControl),, Q#7 survey 3; alcohol last 2 weeks?  !!!
        skippableArray(1, 1) = 25144 'Choice = 'No' !!!
        'skip pattern 3 - survey 1
        skippableArray(2, 0) = 6334   'Q#24 Survey 1 post/Q#23 survey 1 pre; US citizen?  !!!
        skippableArray(2, 1) = 25859  'Choice = 'Yes' !!!
        'skip pattern 6 - Qset Drinking Experiences Q#1
        skippableArray(3, 0) = 10118  'Qset Drinking Experiences Q#1 (pre/post): JS 10/5/09 initially thought that the skip was used only after Survey 1 however, apparently this was mistaken...will now enable skip for after Surveys 1 & 3
        skippableArray(3, 1) = 44620   'Choice = 'No'   'ATHENA=45439/LIVE=44620   !!! 
        'skip pattern 7 - Qset Diary Q#1
        skippableArray(4, 0) = 10473 'Qset Diary Q#1!!!  'athena=10602/live=10473
        skippableArray(4, 1) = 46277  'Choice = 'No' !!!  'athena=46663/live=46277
        'skip pattern 8 - Qset Diary Q#8
        skippableArray(5, 0) = 10483  'Qset Diary Q#8!!! 'athena=10612/live=10483
        skippableArray(5, 1) = 46293  'Choice = 'No' !!! 'athena=46679/live=46293
        'skip pattern 9 - Qset Diary Q#13
        skippableArray(6, 0) = 10490   'Qset Diary Q#13!!! 'athena=10619/live=10490
        skippableArray(6, 1) = 46300  'Choice = 'No' !!! 'athena=46686/live=46300
        'skip pattern 10 - Qset Diary Q#18
        skippableArray(7, 0) = 10497   'Qset Diary Q#18!!! 'athena=10626/live=10497
        skippableArray(7, 1) = 46307  'Choice = 'No' !!! 'athena=46693/live=46307
        'skip pattern 11 - Qset Alcohol Settings Q#1
        skippableArray(8, 0) = 11285  'Qset Alcohol Setting Q#1!!! 'athena=10838/live=11285
        skippableArray(8, 1) = 50122 'Choice = 'No' !!! 'athena=47712/live=50122
        'skip pattern 12 - Qset Alcohol Settings Q#5
        skippableArray(9, 0) = 11327 'Qset Alcohol Setting Q#5!!! 'athena=10880/live= 11327
        skippableArray(9, 1) = 50390 'Choice = 'No' !!! 'athena=47980/live= 50390
        'skip pattern 13 - Qset Alcohol Settings Q#7
        skippableArray(10, 0) = 11340   'Qset Alcohol Setting Q#7!!! 'athena=10893/live= 11340
        skippableArray(10, 1) = 50469  'Choice = 'No' !!! 'athena=48059/live= 50469
        'skip pattern 14 - Qset Energy Drinks Q#2
        skippableArray(11, 0) = 11139   'Qset Energy Drinks Q#2!!! 'athena=10957/live=11139
        skippableArray(11, 1) = 49450   'Choice = 'No' !!! 'athena=48460/live= 49450


        'skip pattern 4:  survey 2... 
        'question #5 and #6 on survey 2 may be skipped based on user's answer to Q5 in Survey 1
        'so, look for question #4/page 3 in survey 2 as the trigger to check  
        ' so, for skipPattern4/Survey 2, as the skippable questions (Q5 and 6 -- page 4) are being loaded then first check to see if the user
        ' answered 'no' to the 'did you drink in the last year' question on survey 1
        'so, hardcode Q4b (last questionID on page 3) in Survey 2 here: 
        intSurvey2SkipQuestion = 2914



        For i = 0 To UBound(arrQuestionArray)    'question array from submitted page
            strQuestion = arrQuestionArray(i)
            'for some reason the arrQuestionArray carries an extra value, an empty string at end of array which causes an error here...for now will use this workaround
            If strQuestion <> "" Then
                intQuestionID = CInt(strQuestion)
            Else
                intQuestionID = 0
            End If
            strQuestionID = "a" & intQuestionID

            y = 1
            For x = 0 To UBound(skippableArray)    'array containing skippable questions
                If intQuestionID = skippableArray(x, 0) Then    'skippable question found on page
                    'need to get the questionchoice from this question to see if it should be skipped
                    If skippableArray(x, 1) = CInt(Request.Form.Item(strQuestionID)) Then    'skip
                        Select Case x       'go to function to determine actual skip pattern
                            Case 0
                                pageNumber = skipNoDrinksPastYear() 'Q#5 survey 1 Pre/Post, Q#4 survey 3, Q#3 survey 1/3 PreControl
                                blnSkipForward = True
                                Exit For
                            Case 1
                                pageNumber = skipNoDrinksPastTwoWeeks() 'Q#8 survey 1(pre/post), Q#4 survey 1/3(preControl),, Q#7 survey 3
                                blnSkipForward = True
                                Exit For
                            Case 2
                                pageNumber = skipUSCitizen() 'Q#24 Survey 1 only
                                blnSkipForward = True
                                Exit For
                            Case 3
                                pageNumber = skipQsetDrinkExp1() 'Q#1, Qset Drinking Exp.(after Survey1), Q#1 Qset Drinking Exp.(after Survey3)
                                blnSkipForward = True
                                Exit For
                            Case 4
                                pageNumber = skipQsetDiary1() 'Q#1 Qset, Diary (after Survey1), Q#1 Qset Diary (after Survey3)
                                blnSkipForward = True
                                Exit For
                            Case 5
                                pageNumber = skipQsetDiary8() 'Q#8 Qset, Diary (after Survey1), Q#8 Qset Diary (after Survey3)
                                blnSkipForward = True
                                Exit For
                            Case 6
                                pageNumber = skipQsetDiary13() 'Q#13 Qset, Diary (after Survey1), Q#13 Qset Diary (after Survey3)
                                blnSkipForward = True
                                Exit For
                            Case 7
                                pageNumber = skipQsetDiary18() 'Q#18 Qset, Diary (after Survey1), Q#18 Qset Diary (after Survey3)
                                blnSkipForward = True
                                Exit For
                            Case 8
                                pageNumber = skipQsetAlSettings1() ''Q#1, Qset Alcohol Settings.(after Survey1):
                                blnSkipForward = True
                                Exit For
                            Case 9
                                pageNumber = skipQsetAlSettings5() ''Q#5, Qset Alcohol Settings.(after Survey1):
                                blnSkipForward = True
                                Exit For
                            Case 10
                                pageNumber = skipQsetAlSettings7() ''Q#7, Qset Alcohol Settings.(after Survey1):
                                blnSkipForward = True
                                Exit For
                            Case 11
                                pageNumber = skipQsetEnrgDrinks2() ''Q#2, Qset Energy Drinks(after Survey1):
                                blnSkipForward = True
                                Exit For
                        End Select
                    End If
                    y = y + 1
                End If
            Next
            If blnSkipForward = True Then
                Exit For    'if skip pattern determined then leave this 
            End If

            'this block only for the skip pattern in PostMat Survey 2
            If pageNumber = 1 And intTestTypeID = 10 Then 'page 4 on survey 2 (both pre and post) contains skippable questions #5 and #6
                'If intQuestionID = intSurvey2SkipQuestion Then 'go to skip function
                pageNumber = skipSurvey2NoDrinksPastYear()
                If blnSkipPattern4 Then
                    blnSkipForward = True
                End If
                'End If
            End If
            ''=========UPENN CONTROL GROUP q#26 SURVEY 3 SKIP PATTERN 5==================================
            '' iterate through questions and trap for q#26 survey 3: hardcoded value= "a" + Q#26 survey 3 ID + "text" + Q#26 questionChoiceID
            '' if user input '0' then skip...anything else, a blank field or anything > 0 the user doesn't skip
            If intQuestionID = intUPennTextQuestion Then    'skippable question found on page
                theAnswerArray = Split(Request.Form.Item(strQuestionID & "_MC"), ",")
                For k = 0 To UBound(theAnswerArray)
                    intQuestionChoice = Trim(theAnswerArray(k))
                    'this builds the string to reference the value for this text questionchoice
                    strTextQuestionID = strQuestionID & "text" & intQuestionChoice
                    If strTextQuestionID = "a7335text30585" Then '!!! hardcoded value= a + Q#26 survey 3 ID + text + Q#26 questionChoiceID
                        strTextAnswer = Trim(Request.Form.Item(strTextQuestionID))
                        If (IsNothing(strTextAnswer)) Or (strTextAnswer = "") Then       'javascript ensures that only valid numbers entered
                            blnSkipForward = False 'dont skip if user didn't enter value
                        Else
                            If CInt(strTextAnswer) = 0 Then 'user entered '0' so skip
                                blnSkipForward = True
                                blnSkipPattern5 = True
                                pageNumber = 12 'skipping from Q#26, pg 9 tp Q#31 pg 12
                            Else
                                blnSkipForward = False
                            End If
                        End If
                    End If
                Next
            End If
            ''===============END UPENN CONTROL GROUP SURVEY 3 SKIP PATTERN 5==========================================================
        Next

        If blnSkipForward = True Then
            checkSkipForward = pageNumber
        Else
            checkSkipForward = pageNumber + 1
        End If

    End Function

    Function checkSurvey2SkipForward() As Integer
        Dim blnSkipForward As Boolean
        'this block only for the skip pattern in PostMat Survey 2
        If (pageNumber = 1 Or pageNumber = 3) And intTestTypeID = 10 Then 'page 4 on survey 2 (both pre and post) contains skippable questions #5 and #6
            'If intQuestionID = intSurvey2SkipQuestion Then 'go to skip function
            pageNumber = skipSurvey2NoDrinksPastYear()
            If blnSkipPattern4 Then
                blnSkipForward = True
            End If
            'End If
        End If
        If blnSkipForward Then
            Return pageNumber
        Else
            Return pageNumber + 1
        End If
    End Function


    Function getSurveyQuestions(ByVal iPageNumber As Integer) As String
        'Dim DRSurvey As DataSet
        Dim strHtml As String
        Dim QuestionTable As DataTable
        Dim QuestionColumn As DataColumn
        Dim QuestionRow As DataRow
        Dim i As Integer = 0
        Dim iTestFormID As Integer
        Dim bAnswered As Boolean = False
        'Dim iErrorCode As Integer = 0
        Dim dvAnswers As DataView
        Dim dtAnswers As DataTable
        Dim bCalQuestion As Boolean = False
        Dim bFemale As Boolean
        If Not IsNothing(Session.Item("SurveyFemale")) Then
            bFemale = Session.Item("SurveyFemale")
        Else
            'Get the gender from classification and save it here.
            bFemale = CBool(getUserClassification(2))
            Session.Item("SurveyFemale") = bFemale
        End If


        iTestFormID = getTestFormID()

        If iPageNumber = 1 Or Request.Form("navFunction") = "back" Then
            'Grab all the saved questions in the survey
            loadSurveyAnswers(iTestFormID)
        End If

        'read the answers from the session
        If Not IsNothing(Session("SurveyAnswers")) Then
            dtAnswers = CType(Session("SurveyAnswers"), DataTable)
            dvAnswers = New DataView(dtAnswers, "", "iQuestionChoiceID", DataViewRowState.CurrentRows)
        Else
            'Survey answers weren't found or session is having issues
            Server.Transfer("/error.aspx?msg=surveyEmpty")
        End If

        QuestionTable = getXMLDT(iTestFormID, iPageNumber)

        If iErrorCode = 1 Then    'File doesn't exist for this page, so create it then read it
            createXMLSurvey(iTestFormID, iPageNumber)
            QuestionTable = getXMLDT(iTestFormID, iPageNumber)
        End If

        answeredSomething = 0
        intDataViewIncrement = 0
        strHtml = ""

        'js 5/21/08:  9.0 randomized questions precede custom questions so custom questions
        ' will not be numbered sequentially at this point...they will start at '1'...as time allows we will build in 
        'logic that will discern whether the user got a randomized Qset or not and then number
        'the Qset questions sequentially as well as the custom questions, if any (possibly with the intIOrder below)
        If iCourseTypeID = 1 Then 'postmat
            intCustomQuestionNumber = 0 'see above comments
        Else 'premat
            intCustomQuestionNumber = 0 'see above comments
        End If



        strHtml = strHtml & "<table width=""100%"" cellspacing=0 cellpadding=1 border=0>"

        'all the code for setting up and formatting the questions will go here
        For Each QuestionRow In QuestionTable.Rows
            'loop through all survey questions
            intIOrder = QuestionRow.Item("iOrder")
            If intIOrder > 5000 Then    'custom questions all start with 5000/Qset questions all start with 6000
                blnCustomQuestion = True
            End If
            intQuestionID = QuestionRow.Item("iQuestionID")
            intQuestionFormatID = QuestionRow.Item("iQuestionFormatID")
            intQuestionTypeID = QuestionRow.Item("iQuestionTypeID")
            intQuestionNumber = QuestionRow.Item("iQuestionNumber")
            intlastPage = QuestionRow.Item("iLastPage")
            intPageStart = QuestionRow.Item("iPageStart")
            bAnswered = False

            'if we are coming out of a questionchoice sub-loop then check to
            ' make sure we aren't on the same questionID...compare the
            ' DATAVIEW questionID (intHoldDataViewQuestionID) to the DATASET questionID (intQuestionID)
            ' and if the same then continue to skip to go to the next new questionID...
            If intHoldDataViewQuestionID <> intQuestionID Then
                If CInt(QuestionRow.Item("tiSubQuestion")) = 0 Then
                    If intQuestionFormatID <> 1 And intQuestionFormatID <> 8 Then
                        strHtml = strHtml & "<tr><td>&nbsp;</td></tr>"
                    End If
                    If intQuestionTypeID <> 14 And intQuestionTypeID <> 15 And intQuestionTypeID <> 13 And intQuestionTypeID <> 60 Then
                        If CBool(QuestionRow.Item("tiDisplayQuestionNumber")) Then
                            intCustomQuestionNumber = intCustomQuestionNumber + 1
                            strHtml = strHtml & "<tr>" & vbCrLf
                            strHtml = strHtml & "<td class=""questionText"" style=""questionText""><b>"
                            If blnCustomQuestion And intIOrder < 6000 Then       'its a custom question (and not a Qset question) so we need to use dynamic numbering
                                strHtml = strHtml & intCustomQuestionNumber & ".&nbsp;&nbsp;</b>"
                            Else
                                strHtml = strHtml & QuestionRow.Item("iQuestionNumber") & ".&nbsp;&nbsp;</b>"
                            End If
                            If intQuestionTypeID = 10 Then
                                If InStr(QuestionRow.Item("txQuestion"), "<<DrinkQTY>>") Then
                                    strHtml = strHtml & Replace(QuestionRow.Item("txQuestion"), "<<DrinkQTY>>", getHighestDrinkValue())
                                Else
                                    strHtml = strHtml & QuestionRow.Item("txQuestion")
                                End If
                            Else
                                If intQuestionID = 6246 Or intQuestionID = 6256 Then
                                    'These have different text if female.
                                    strHtml = strHtml & editForFemale(QuestionRow.Item("txQuestion"), bFemale)
                                Else
                                    strHtml = strHtml & QuestionRow.Item("txQuestion")
                                End If

                            End If
                            strHtml = strHtml & "</td>" & vbCrLf
                            strHtml = strHtml & "</tr>" & vbCrLf
                        Else
                            strHtml = strHtml & "<tr>" & vbCrLf
                            strHtml = strHtml & "<td class=""questionText"" style=""questionText"">"
                            strHtml = strHtml & QuestionRow.Item("txQuestion")
                            strHtml = strHtml & "</td>" & vbCrLf
                            strHtml = strHtml & "</tr>" & vbCrLf
                        End If
                    End If

                    If CBool(QuestionRow.Item("tiAnswerable")) Then
                        Session.Item("skipSave") = False
                        strQuestions = strQuestions & intQuestionID & ","
                        If intQuestionTypeID <> 16 Then
                            strHtml = strHtml & "<tr><td>"
                        End If
                        'Comment Key:	MCSR - Multiple Choice Single Response
                        '				MCMR - Multiple Choice Multiple Response
                        '				MFR - Multiple Free Response

                        Select Case intQuestionTypeID       'Type of question
                            Case 18     'Calendar Question
                                bCalQuestion = True
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                Dim strHiddenFieldValues As String
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                'intDataViewIncrement = intDataViewIncrement + 1
                                intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                If intDataViewQuestionID = intQuestionID Then
                                    'intDataViewIncrement = intDataViewIncrement + 1
                                    'Think we incriment this in the calendar grid function
                                    intHoldDataViewQuestionID = intDataViewQuestionID
                                    'Call the calendar grid
                                    strHtml = strHtml & createCalendarGrid(QuestionChoiceDataView, dvAnswers, intQuestionID, intDataViewIncrement, intHoldDataViewQuestionID, strReqQuestions, strTextQuestion, strHiddenFieldValues)
                                    'create the hidden form fields
                                    strHtml = strHtml & "<input type=""hidden"" name=""a" & intQuestionID & "_MC"" value=""" & strHiddenFieldValues & """>" & vbCrLf
                                    'put a hidden form field so we can tell when we're saving answers that we need to store these in the session too.
                                    strHtml = strHtml & "<input type=""hidden"" name=""calQuestion"" value=""" & intQuestionID & """>" & vbCrLf
                                    'Move on
                                End If

                            Case 6, 1, 3          'Vertical, MCSR, MCMR, MFR
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                strHtml = strHtml & "<table "
                                If intQuestionFormatID = 1 Then
                                    strHtml = strHtml & "cellpadding=0 cellspacing=0"
                                End If
                                strHtml = strHtml & ">" & Chr(13)
                                'loop to display question choices
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    bAnswered = False
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        'Now check for this answer in the dvAnswers DataView
                                        If intQuestionTypeID = 6 Then             'Only allow the Open Text questsions to use this one.
                                            responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                            'HRH 2/7/07 Adding this to not show pop-up if the free response is the only question on the page
                                            If QuestionChoiceDataView.Count = 1 Then
                                                'There is only one question on the page
                                                answeredSomething = 1
                                            End If
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 1")
                                        End If
                                        'If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                        'FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                        'responseSelected = True
                                        'Else
                                        'FreeResponse = ""
                                        'responseSelected = False
                                        'End If
                                        strHtml = strHtml & "<tr>"

                                        If intQuestionFormatID = 4 Then             ' LARGE Question (or Text Areas)
                                            strHtml = strHtml & "<td class=answerNumber><label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                            strHtml = strHtml & "<textarea "
                                            strHtml = strHtml & "onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
                                            strHtml = strHtml & " name=""a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                            strHtml = strHtml & " rows=5 cols=40 onKeyUp=""return maxCharaters(this)"">" & FreeResponse & "</textarea>" & Chr(13)
                                            strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """>" & Chr(13)
                                            strTextQuestion = strTextQuestion & "," & intQuestionID
                                        Else
                                            strHtml = strHtml & "<td valign=top width=25 class=""answerNumber1"" style=""answerNumber1""><label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                            strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();""  type="

                                            Select Case intQuestionTypeID
                                                Case 1
                                                    strHtml = strHtml & "radio"
                                                Case 3
                                                    strHtml = strHtml & "checkbox "
                                                    If responseSelected Then
                                                        strHtml = strHtml & " checked"
                                                        answeredSomething = 1
                                                    End If

                                                Case 6
                                                    strHtml = strHtml & "text size=3"
                                                    If QuestionChoiceDataView.Item(i).Row("iQuestionResponseTypeID") = 1 Then
                                                        intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                                        If Not IsDBNull(intMax) Then                'its a question that needs specific min/max validation
                                                            strHtml = strHtml & " maxlength=3 onBlur=""validateMinMax(this," & QuestionChoiceDataView.Item(i).Row("iMinimum") & ", " & intMax & ")"""
                                                        Else
                                                            strHtml = strHtml & " maxlength=3 onBlur=""validate(this)"""
                                                        End If
                                                    Else
                                                        strHtml = strHtml & " maxlength=500"
                                                    End If
                                            End Select
                                            'I commented this out because this function was being called 2x for each question choice. 
                                            bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 2")
                                            If bAnswered And intQuestionTypeID <> 22 Then
                                                strHtml = strHtml & " checked"
                                                answeredSomething = 1
                                                'Else
                                                '	answeredSomething = 0
                                            End If
                                            'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") And intQuestionTypeID <> 22 Then
                                            'strHtml = strHtml & " checked"
                                            'End If
                                            'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") And intQuestionTypeID <> 22 Then
                                            '	answeredSomething = 1
                                            'End If
                                            strHtml = strHtml & " name=""a" & intQuestionID
                                            If intQuestionTypeID = 3 Then
                                                strHtml = strHtml & "_MC"
                                            End If
                                            If intQuestionTypeID <> 6 Then
                                                strHtml = strHtml & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">"
                                            Else
                                                strHtml = strHtml & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " value=""" & QuestionChoiceDataView.Item(i).Row("vchFreeResponse") & """ >" & Chr(13)
                                                If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" And Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchFreeResponse")) Then
                                                    answeredSomething = 1
                                                End If 'HRH Added MC to below hidden field 4/1/08
                                                strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & "_MC"" value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """>" & Chr(13)
                                                strTextQuestion = strTextQuestion & "," & intQuestionID
                                            End If
                                        End If
                                        If (intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And intQuestionTypeID <> 6 Then
                                            Trace.Write("QuestionFormatInfo", CStr(iQuestionID) & " " & QuestionRow.Item(17))
                                            strHtml = strHtml & QuestionRow.Item("chSelectionNumber")
                                        End If
                                        If Not (IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice"))) Then
                                            strHtml = strHtml & "</label></td><td valign=top class=""answerText"" style=""answerText""><label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">"
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "</label>"
                                        End If
                                        If QuestionChoiceDataView.Item(i).Row("iSubQuestionID") <> 0 Then
                                            strSubQuestions = getSubQuestion(QuestionChoiceDataView.Item(i).Row("iSubQuestionID"), QuestionTable.Copy)
                                            strHtml = strHtml & strSubQuestions
                                            'will need to increment for the sub-question
                                            strSubQuestions = ""
                                            intDataViewIncrement = intDataViewIncrement + 1
                                        End If
                                        strHtml = strHtml & "</label></td></tr>" & Chr(13)
                                    Else
                                        'finish table before exiting...
                                        strHtml = strHtml & "</table>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 2, 4, 7          ' Horizontal, MCSR, MCMR, MFR
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                strHtml = strHtml & "<table width=""100%"" "
                                If intQuestionFormatID = 1 Then
                                    strHtml = strHtml & "cellpadding=0 cellspacing=0"
                                End If
                                strHtml = strHtml & ">" & Chr(13)
                                strHtml = strHtml & "<tr>"
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        'Now check for this answer in the dvAnswers DataView
                                        responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                        'Code to track how often this is called.
                                        'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 3")
                                        'If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                        'FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                        'responseSelected = True
                                        'Else
                                        'FreeResponse = ""
                                        'responseSelected = False
                                        'End If
                                        strHtml = strHtml & "<td valign=bottom width=""10%"" class=""answerNumber"" style=""answerNumber""><label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        If (intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7) And (Not (IsNothing(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")))) Then
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>" & Chr(13)
                                        End If
                                        If intQuestionFormatID = 4 Then             ' LARGE Question (or Text Areas)
                                            strHtml = strHtml & "<textarea "
                                            strHtml = strHtml & "onChange=""setAnsweredSomething();"""
                                            strHtml = strHtml & " name=""a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")              '1/4/01 JC
                                            strHtml = strHtml & """ rows=5 cols=40 onKeyUp=""return maxCharaters(this)"">" & FreeResponse & "</textarea>" & Chr(13)             '1/4/00: jc add freeresponse
                                            strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID
                                            If intQuestionTypeID = 20 Then
                                                strHtml = strHtml & "_MC"
                                            End If
                                            strHtml = strHtml & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)            '1/4/00: jc to get when saving
                                            strTextQuestion = strTextQuestion & "," & intQuestionID
                                        Else
                                            strHtml = strHtml & "<input "
                                            strHtml = strHtml & "onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
                                            strHtml = strHtml & " type="
                                            Select Case intQuestionTypeID
                                                Case 2
                                                    strHtml = strHtml & "radio"
                                                Case 4
                                                    strHtml = strHtml & "checkbox"
                                                    If responseSelected Then
                                                        strHtml = strHtml & " checked"
                                                        answeredSomething = 1
                                                    End If

                                                Case 7
                                                    strHtml = strHtml & "text size=2 "
                                                    If QuestionChoiceDataView.Item(i).Row("iQuestionResponseTypeID") = 1 Then                '--NOTE: This was iFreeResponseID												 '1/13/01 MB
                                                        intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                                        If Not IsDBNull(intMax) Then                'its a question that needs specific min/max validation
                                                            strHtml = strHtml & " maxlength=2 onBlur=""validateMinMax(this," & QuestionChoiceDataView.Item(i).Row("iMinimum") & ", " & intMax & ")"""
                                                        Else
                                                            strHtml = strHtml & " maxlength=2 onBlur=""validate(this)"""
                                                        End If
                                                        If strReqQuestions <> "" Then
                                                            strReqQuestions = strReqQuestions & ",a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                                        Else
                                                            strReqQuestions = "a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                                        End If
                                                    Else
                                                        strHtml = strHtml & " maxlength=500"
                                                    End If
                                            End Select
                                            bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 4")
                                            If bAnswered And intQuestionTypeID <> 22 Then
                                                strHtml = strHtml & " checked"
                                                answeredSomething = 1
                                                'Else
                                                '	answeredSomething = 0
                                            End If
                                            'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") Then
                                            '	strHtml = strHtml & " checked"
                                            'End If
                                            'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") And intQuestionTypeID <> 23 Then
                                            '	answeredSomething = 1
                                            'End If
                                            strHtml = strHtml & " name=""a" & intQuestionID
                                            If intQuestionTypeID = 4 Then
                                                strHtml = strHtml & "_MC)"
                                            End If
                                            If intQuestionTypeID <> 7 Then
                                                strHtml = strHtml & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Else
                                                strHtml = strHtml & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " value=""" & FreeResponse & """  >" & Chr(13)
                                                If FreeResponse <> "" And Not IsDBNull(FreeResponse) Then
                                                    answeredSomething = 1
                                                End If
                                                strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & "_MC" & """ value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                                strTextQuestion = strTextQuestion & "," & intQuestionID
                                            End If
                                        End If
                                        If intQuestionFormatID <> 3 Then
                                            If (intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And intQuestionTypeID <> 7 Then
                                                Trace.Write("QuestionFormatInfo", CStr(iQuestionID) & " " & QuestionRow.Item(17))
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>" & Chr(13)
                                            End If
                                        Else
                                            strHtml = strHtml & "</label></td><td class=""answerText"" style=""answerText"">" & Chr(13)
                                        End If
                                        If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")) Then

                                            If Not ((intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7)) Then
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "</label>" & Chr(13)
                                            End If
                                        End If

                                        If QuestionChoiceDataView.Item(i).Row("iSubQuestionID") <> 0 Then
                                            Call getSubQuestion(QuestionChoiceDataView.Item(i).Row("iSubQuestionID"), QuestionTable.Copy)

                                        End If
                                        strHtml = strHtml & "</td>" & Chr(13)
                                    Else
                                        'finish table before exiting...
                                        strHtml = strHtml & "</tr>" & Chr(13)
                                        strHtml = strHtml & "</table>" & Chr(13)
                                        strHtml = strHtml & "<tr><td>&nbsp;</td></tr>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 17       ' Horizontal Seven Scale
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                strHtml = strHtml & "<table width=""100%"" "
                                If intQuestionFormatID = 1 Then
                                    strHtml = strHtml & "cellpadding=0 cellspacing=0"
                                End If
                                strHtml = strHtml & ">" & Chr(13)
                                strHtml = strHtml & "<tr>"
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        'Now check for this answer in the dvAnswers DataView
                                        responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                        strHtml = strHtml & "<td valign=bottom width=""10%"" class=""answerNumber"" style=""answerNumber""><label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        If (intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7) And (Not (IsNothing(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")))) Then
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>" & Chr(13)
                                        End If
                                        If intQuestionFormatID = 4 Then             ' LARGE Question (or Text Areas)
                                            strHtml = strHtml & "<textarea "
                                            strHtml = strHtml & "onChange=""setAnsweredSomething();"""
                                            strHtml = strHtml & " name=""a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")              '1/4/01 JC
                                            strHtml = strHtml & """ rows=5 cols=40 onKeyUp=""return maxCharaters(this)"">" & FreeResponse & "</textarea>" & Chr(13)             '1/4/00: jc add freeresponse
                                            strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID
                                            If intQuestionTypeID = 20 Then
                                                strHtml = strHtml & "_MC"
                                            End If
                                            strHtml = strHtml & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)            '1/4/00: jc to get when saving
                                            strTextQuestion = strTextQuestion & "," & intQuestionID
                                        Else
                                            strHtml = strHtml & "<input "
                                            strHtml = strHtml & "onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
                                            strHtml = strHtml & " type="
                                            Select Case intQuestionTypeID
                                                Case 17
                                                    strHtml = strHtml & "radio"
                                                Case 4
                                                    strHtml = strHtml & "checkbox"
                                                    If responseSelected Then
                                                        strHtml = strHtml & " checked"
                                                        answeredSomething = 1
                                                    End If

                                                Case 7
                                                    strHtml = strHtml & "text size=2 "
                                                    If QuestionChoiceDataView.Item(i).Row("iQuestionResponseTypeID") = 1 Then                '--NOTE: This was iFreeResponseID												 '1/13/01 MB
                                                        intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                                        If Not IsDBNull(intMax) Then                'its a question that needs specific min/max validation
                                                            strHtml = strHtml & " maxlength=2 onBlur=""validateMinMax(this," & QuestionChoiceDataView.Item(i).Row("iMinimum") & ", " & intMax & ")"""
                                                        Else
                                                            strHtml = strHtml & " maxlength=2 onBlur=""validate(this)"""
                                                        End If
                                                        If strReqQuestions <> "" Then
                                                            strReqQuestions = strReqQuestions & ",a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                                        Else
                                                            strReqQuestions = "a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                                        End If
                                                    Else
                                                        strHtml = strHtml & " maxlength=500"
                                                    End If
                                            End Select
                                            bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 4")
                                            If bAnswered And intQuestionTypeID <> 22 Then
                                                strHtml = strHtml & " checked"
                                                answeredSomething = 1
                                                'Else
                                                '	answeredSomething = 0
                                            End If
                                            strHtml = strHtml & " name=""a" & intQuestionID
                                            If intQuestionTypeID = 4 Then
                                                strHtml = strHtml & "_MC)"
                                            End If
                                            If intQuestionTypeID <> 7 Then
                                                strHtml = strHtml & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Else
                                                strHtml = strHtml & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " value=""" & FreeResponse & """  >" & Chr(13)
                                                If FreeResponse <> "" And Not IsDBNull(FreeResponse) Then
                                                    answeredSomething = 1
                                                End If
                                                strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & "_MC" & """ value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                                strTextQuestion = strTextQuestion & "," & intQuestionID
                                            End If
                                        End If
                                        If intQuestionFormatID <> 3 Then
                                            If (intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And intQuestionTypeID <> 7 Then
                                                Trace.Write("QuestionFormatInfo", CStr(iQuestionID) & " " & QuestionRow.Item(17))
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>" & Chr(13)
                                            End If
                                        Else
                                            strHtml = strHtml & "</label></td><td class=""answerText"" style=""answerText"">" & Chr(13)
                                        End If
                                        If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")) Then

                                            If Not ((intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7)) Then
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "</label>" & Chr(13)
                                            End If
                                        End If

                                        If QuestionChoiceDataView.Item(i).Row("iSubQuestionID") <> 0 Then
                                            Call getSubQuestion(QuestionChoiceDataView.Item(i).Row("iSubQuestionID"), QuestionTable.Copy)

                                        End If
                                        strHtml = strHtml & "</td>" & Chr(13)
                                    Else
                                        'finish table before exiting...
                                        strHtml = strHtml & "</tr>" & Chr(13)
                                        strHtml = strHtml & "</table>" & Chr(13)
                                        strHtml = strHtml & "<tr><td>&nbsp;</td></tr>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 10          '  Fill in the blank
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        'Now check for this answer in the dvAnswers DataView
                                        responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                        'Code to track how often this is called.
                                        'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 5")
                                        'If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                        'FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                        'responseSelected = True
                                        'Else
                                        'FreeResponse = ""
                                        'responseSelected = False
                                        'End If
                                        strHtml = strHtml & "<label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        strHtml = strHtml & "<input "
                                        strHtml = strHtml & "onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
                                        strHtml = strHtml & " type=text size=3 "
                                        If QuestionChoiceDataView.Item(i).Row("iQuestionResponseTypeID") = 1 Then             'This was iFreeResponseID
                                            intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                            If Not IsDBNull(intMax) Then             'its a question that needs specific min/max validation
                                                strHtml = strHtml & " maxlength=3 onBlur=""validateMinMax(this," & QuestionChoiceDataView.Item(i).Row("iMinimum") & ", " & intMax & ")"""
                                            Else
                                                strHtml = strHtml & " maxlength=3 onBlur=""validate(this)"""
                                            End If
                                            If strReqQuestions <> "" Then
                                                strReqQuestions = strReqQuestions & ",a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                            Else
                                                strReqQuestions = "a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                            End If
                                        Else
                                            strHtml = strHtml & " maxlength=500"
                                        End If
                                        '===attempt to insert onFocus event into td tag so that if its the weight question, use the onFocus to allow user input only into EITHER 'pounds' or 'kilograms' box====================
                                        Dim strWeightChoice As String
                                        strWeightChoice = "a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                        If intQuestionID = 6338 Then 'the weight question
                                            If strWeightChoice = "a6338text25877" Then 'pounds box so erase kilogram box
                                                strHtml = strHtml & " onFocus=""document.theForm.a6338text25878.value='';  """
                                            ElseIf strWeightChoice = "a6338text25878" Then 'kilogram box so erase pounds box
                                                strHtml = strHtml & " onFocus=""document.theForm.a6338text25877.value='';  """
                                            End If
                                        End If
                                        '=============================================================================================================================================
                                        strHtml = strHtml & " name=""a" & intQuestionID
                                        strHtml = strHtml & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ value=""" & FreeResponse & """ >"

                                        If QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") <> "" Then
                                            strHtml = strHtml & "&nbsp;" & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & Chr(13)
                                        End If
                                        If FreeResponse <> "" And Not IsDBNull(FreeResponse) Then
                                            answeredSomething = 1
                                        End If
                                        strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & "_MC" & """ value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                        strTextQuestion = strTextQuestion & "," & intQuestionID
                                        strHtml = strHtml & "</label>&nbsp; &nbsp;" & Chr(13)
                                    Else
                                        'finish row before exiting...
                                        strHtml = strHtml & "<tr><td>&nbsp;</td></tr>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 14          'Tight Row

                                'js 4/14/09: for the new grid-style-checkbox...
                                'let's start by hardcoding in a new questionformatID (before creating a new one in the DB):
                                'intQuestionFormatID = 18

                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                intChoicesCount = 0
                                j = j + 1
                                'js 4/14/09: for the new grid-style-checkbox...intQuestionFormatID = 18
                                If intQuestionFormatID = 1 Or intQuestionFormatID = 17 Or intQuestionFormatID = 18 And bSpecialFormat <> True Then
                                    strHtml = strHtml & "<table width=780 cellpadding=0 cellspacing=0 border=0>" & Chr(13)
                                End If
                                strStyle = "surveyQuestionText"
                                If CBool(j Mod 2) Then
                                    strStyle = strStyle & "1"
                                End If
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)

                                strHtml = strHtml & "<tr><td class=""" & strStyle & """ style=""" & strStyle & """"
                                strHtml = strHtml & " width=" & QuestionChoiceDataView.Item(i).Row("iQColumnWidth") & ">"
                                If intQuestionID = 6246 Or intQuestionID = 6256 Then
                                    'These have different text if female.
                                    strHtml = strHtml & editForFemale(QuestionChoiceDataView.Item(i).Row("txQuestion"), bFemale) & Chr(13)
                                Else
                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                End If
                                'strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                strHtml = strHtml & "</td>"
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")

                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        intChoicesCount = intChoicesCount + 1
                                        strHtml = strHtml & "<td width=" & QuestionChoiceDataView.Item(i).Row("iColumnWidth") & " class=""" & strStyle & """ style=""" & strStyle & """"
                                        strHtml = strHtml & "><label for=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)

                                        If intQuestionFormatID = 17 Then 'for dropdown
                                            strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                            'js 4/14/09: for the new grid-style-checkbox...
                                            'ElseIf intQuestionFormatID = 18 Then 'for checkboxes
                                            'strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=checkbox"
                                        Else 'for radio buttons or anything not above
                                            strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                        End If
                                        If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) Then
                                            bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                        Else
                                            bAnswered = False
                                        End If

                                        If bAnswered Then
                                            If intQuestionFormatID <> 17 Then
                                                strHtml = strHtml & " checked"
                                            End If
                                            answeredSomething = 1
                                        End If

                                        If intQuestionFormatID = 17 Then 'for dropdown
                                            strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Dim a As Integer
                                            Dim strSelected As String
                                            a = 0
                                            For a = a To 14
                                                If vchDropDown = CStr(a) Then
                                                    strSelected = "selected"
                                                Else
                                                    strSelected = ""
                                                End If
                                                strHtml = strHtml & "<option " & strSelected & ">" & a & "</option>"
                                            Next

                                            strHtml = strHtml & "</select>"
                                            strHtml = strHtml & "</label></td>" & Chr(13)
                                            strHtml = strHtml & "<input type=hidden name=dropdown" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                        Else 'any format not dropdown...radio or checkbox
                                            strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></label></td>" & Chr(13)
                                        End If

                                        If QuestionChoiceDataView.Item(i).Row("iQuestionID") <> intQuestionID Then
                                            If intChoicesCount < intNoChoices Then
                                                strHtml = strHtml & "<td width=" & QuestionChoiceDataView.Item(i).Row("iColumnWidth") & " class=""" & strStyle & """ style=""" & strStyle & """>&nbsp;</td>" & Chr(13)
                                            End If
                                        End If
                                    Else
                                        'finish table before exiting...
                                        If intChoicesCount < intNoChoices Then
                                            strHtml = strHtml & "<td class=""" & strStyle & """ style=""" & strStyle & """>&nbsp;</td>" & Chr(13)
                                        End If
                                        strHtml = strHtml & "</tr>" & Chr(13)

                                        'for radio, checkbox, and dropdown formats
                                        'js 4/14/09: for the new grid-style-checkbox...
                                        If intQuestionFormatID = 1 Or intQuestionFormatID = 17 Or intQuestionFormatID = 18 Then
                                            strHtml = strHtml & "</table>" & Chr(13)
                                        End If
                                        Exit For
                                    End If
                                Next

                            Case 15          'Tight Row Header
                                intNoChoices = 0
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)

                                'js 4/14/09: for the new grid-style-checkbox...
                                'let's start by hardcoding in a new questionformatID (before creating a new one in the DB):
                                'intQuestionFormatID = 18

                                j = 1
                                strHtml = strHtml & "<table width=780 cellpadding=0 cellspacing=0 border=0>" & Chr(13)
                                strStyle = "surveyQuestionText"
                                If CBool(j Mod 2) Then
                                    strStyle = strStyle & "1"
                                End If
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)

                                strHtml = strHtml & "<tr><td class=""" & strStyle & """ style=""" & strStyle & """"
                                strHtml = strHtml & " width=" & QuestionChoiceDataView.Item(i).Row("iQColumnWidth") & ">"
                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                strHtml = strHtml & "&nbsp;&nbsp;</td>"
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        intIdcounter = intIdcounter + 1
                                        intNoChoices = intNoChoices + 1

                                        strHtml = strHtml & "<td width=" & QuestionChoiceDataView.Item(i).Row("iColumnWidth") & " class=""" & strStyle & """ style=""" & strStyle & """"
                                        strHtml = strHtml & "><label for=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        If Not (IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice"))) Then
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>"
                                        Else
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & " < br > """
                                        End If

                                        If intQuestionFormatID = 17 Then 'for dropdown
                                            strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                            'ElseIf intQuestionFormatID = 18 Then 'for checkbox
                                            'strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=checkbox"
                                        Else 'for radio + any format not above
                                            strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                        End If

                                        bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                        If bAnswered Then
                                            If intQuestionFormatID <> 17 Then
                                                strHtml = strHtml & " checked"
                                            End If
                                            answeredSomething = 1
                                        End If

                                        If intQuestionFormatID = 17 Then
                                            strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Dim a As Integer
                                            Dim strSelected As String
                                            a = 0
                                            For a = a To 14
                                                If vchDropDown = CStr(a) Then
                                                    strSelected = "selected"
                                                Else
                                                    strSelected = ""
                                                End If
                                                strHtml = strHtml & "<option " & strSelected & ">" & a & "</option>"
                                            Next

                                            strHtml = strHtml & "</select>"
                                            strHtml = strHtml & "</label></td>" & Chr(13)
                                            strHtml = strHtml & "<input type=hidden name=dropdown" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                        Else
                                            strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></label></td>" & Chr(13)
                                        End If
                                    Else
                                        'finish table before exiting...
                                        strHtml = strHtml & "</tr>" & Chr(13)
                                        strHtml = strHtml & "</table>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 19  'Standard DropDown
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        'Now check for this answer in the dvAnswers DataView
                                        responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                        strHtml = strHtml & "<label for=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                        bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                        If bAnswered Then
                                            answeredSomething = 1
                                        End If
                                        strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                        Dim a As Integer

                                        '=============NEW CODE FOR HANDLING MULTIPLE DROPDOWNS WITH A SINGLE SURVEY=====================================
                                        'JS:all code between '===' is new code added 6/14 to handle multiple drop downs within one survey 
                                        'will proceed with the concept of using the hardcoded questionIDs of the drop down type questions...not the best way to
                                        '  go but have no time to do it more elegantly
                                        Dim arrDropDown() As String
                                        Dim strTypeArray As String

                                        Select Case intQuestionID
                                            Case 6339 'state dropdown  Pre/Post Survey 1 (6339 is same QuestionID for Athena + Live) !!!
                                                strTypeArray = "states"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10053 'Time Spent' Qset - question 1 (GPA)  10061=ATHENA ; 10053=LIVE  !!!
                                                strTypeArray = "gpa"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10087, 10096, 10098, 10100, 10102, 10104, 10106  'Time Spent' Qset - question 26, 29, 31, 33, 35, 37, 39 ATHENA  !!!
                                                'Case 10087, 10096, 10098, 10100, 10102, 10104, 10106  'Time Spent' Qset - question 26, 29, 31, 33, 35, 37, 39 LIVE  !!!
                                                strTypeArray = "time"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10149, 10150, 10151, 10152, 10153, 10154, 10155, 10156, 10157, 10158, 10159 'ATHENA=10265, 10266, 10267, 10268, 10269, 10270, 10271, 10272, 10273, 10274, 10275 'DrinkExp' Qset - question 17 - 27 ATHENA; LIVE=10149, 10150, 10151, 10152, 10153, 10154, 10155, 10156, 10157, 10158, 10159  !!!
                                                'Case  'DrinkExp' Qset - question 17 - 27 LIVE SITE
                                                strTypeArray = "drinks"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10481, 10488, 10495, 10502, 11175, 11177  'athena = 10610, 10617, 10624, 10631 'Diary' Qset - questions 6, 11, 16, 21   ATHENA/ live site= 10481, 10488, 10495, 10502  !!!
                                                strTypeArray = "drinks" 'this block of questionIDs is from the DIARY QSET -- I'm breaking out onto different lines all of the question IDs that use the 'number of drinks' drop down
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10993, 10995  'athena =  'Energy Drinks' Qset - questions 5,7   ATHENA=10993,10995/ live site=   !!!
                                                strTypeArray = "drinks" 'this block of questionIDs is from the ENERGY DRINKS QSET -- I'm breaking out onto different lines all of the question IDs that use the 'number of drinks' drop down
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10140, 10141, 10142, 10143, 10144, 10145, 10146, 10147, 10148 'athena=10256, 10257, 10258, 10259, 10260, 10261, 10262, 10263, 10264 ' Drinking Exp.' Qset, question 8 - 16 ATHENA/live = 10140,10141,10142,10143,10144,10145,10146,10147,10148  !!!
                                                strTypeArray = "consume"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10475 ' Diary Qset, question 2 ATHENA=10604/live=10475  !!!
                                                strTypeArray = "recentday"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10478, 10485, 10492, 10499 'athena=10478,10485,10492,10499 10607, 10614, 10621, 10628 ' Diary Qset, question 5, 10,15, 20 ATHENA/live site= 10478,10485,10492,10499  !!!
                                                strTypeArray = "location"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10477, 10482, 10484, 10489, 10491, 10496, 10498, 10503  'athena=10606, 10611, 10613, 10618, 10620, 10625, 10627, 10632 ' Diary Qset, question 4,7,9,12,14,17,19,22 ATHENA/live site=10477,10482,10484,10489,10491,10496,10498,10503  !!!
                                                strTypeArray = "hours"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10990, 10994, 11172, 11176  ' Energy Drinks Qset, question 4,6 athena=10990,10994 /live site=  !!!
                                                strTypeArray = "days"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case 10993, 10995  ' Energy Drinks Qset, question 5,7 athena=10993,10995 /live site=  !!!
                                                strTypeArray = "nrgdrinks"
                                                arrDropDown = makeArray(strTypeArray)
                                            Case Else
                                                strTypeArray = "error"
                                                arrDropDown = makeArray(strTypeArray)
                                        End Select

                                        Dim strSelected As String
                                        a = 0
                                        For a = a To UBound(arrDropDown)
                                            If vchDropDown = CStr(arrDropDown(a)) Then
                                                strSelected = "selected"
                                            Else
                                                strSelected = ""
                                            End If
                                            strHtml = strHtml & "<option " & strSelected & ">" & arrDropDown(a) & "</option>"
                                        Next

                                        'JS:all code between '===' is new code added 6/14 to handle multiple drop downs with one survey 
                                        '=============NEW CODE FOR HANDLING MULTIPLE DROPDOWNS WITH A SINGLE SURVEY=====================================

                                        '==========================ORIGINAL CODE FOR 'STATES' DROPDOWN=====================================================
                                        'Dim arrStates() As String
                                        '' to do: as time allows, modularize this so that we can call a function here that will load whatever array is needed, ie, if there is more than one drop down that uses an array in the survey...
                                        'arrStates = New String() {"Choose a State", "Alabama", "Alaska", "American Samoa", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "District Of Columbia", "Federated States of Micronesia", "Florida", "Georgia", "Guam", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Marshall Islands", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Northern Mariana Islands", "Ohio", "Oklahoma", "Oregon", "Palau", "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virgin Islands", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", "Did not graduate in the United States"}

                                        'Dim strSelected As String
                                        'a = 0
                                        'For a = a To UBound(arrStates)
                                        '    If vchDropDown = CStr(arrStates(a)) Then
                                        '        strSelected = "selected"
                                        '    Else
                                        '        strSelected = ""
                                        '    End If
                                        '    strHtml = strHtml & "<option " & strSelected & ">" & arrStates(a) & "</option>"
                                        'Next
                                        '===========================ORIGINAL CODE FOR 'STATES' DROPDOWN===================================================

                                        strHtml = strHtml & "</select>"
                                        strHtml = strHtml & "</label></td>" & Chr(13)
                                        strHtml = strHtml & "<input type=hidden name=dropdown" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                    Else
                                        'finish row before exiting...
                                        strHtml = strHtml & "<tr><td>&nbsp;</td></tr>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 13          'Special Header for yes/no/not applicable questions
                                intNoChoices = 0
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                bSpecialFormat = True
                                j = j + 1
                                strHtml = strHtml & "<table width=780 cellpadding=0 cellspacing=0 border=0>" & Chr(13)
                                strStyle = "surveyQuestionText"
                                If CBool(j Mod 2) Then
                                    strStyle = strStyle & "1"
                                End If
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)

                                strHtml = strHtml & "<tr><td class=""" & strStyle & """ style=""" & strStyle & """"
                                strHtml = strHtml & " width=" & QuestionChoiceDataView.Item(i).Row("iQColumnWidth") & ">"
                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                strHtml = strHtml & "</td>"

                                For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")

                                    If intDataViewQuestionID = intQuestionID Then
                                        intDataViewIncrement = intDataViewIncrement + 1
                                        intHoldDataViewQuestionID = intDataViewQuestionID
                                        intNoChoices = intNoChoices + 1
                                        intIdcounter = intIdcounter + 1
                                        strHtml = strHtml & "<td width=""60"" class=""" & strStyle & """ style=""" & strStyle & """"
                                        strHtml = strHtml & "><label for=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " accesskey=" & strAccessKey & ">" & Chr(13)
                                        If Trim(QuestionChoiceDataView.Item(i).Row("chSelectionNumber")) <> "" Then
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>"
                                        Else
                                            strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>"
                                        End If
                                        strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                        bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                        'Code to track how often this is called.
                                        'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 8")
                                        If bAnswered Then
                                            strHtml = strHtml & " checked"
                                            answeredSomething = 1
                                        End If
                                        'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") Then
                                        'strHtml = strHtml & " checked"
                                        'answeredSomething = 1
                                        'End If

                                        strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></label></td>" & Chr(13)
                                    Else
                                        'finish table before exiting...
                                        'and write the 'not applicable' field if it exists as a question choice
                                        '** WILL HAVE TO CHANGE THE INTQUESTIONFORMATID BELOW IF DIFFERENT ON ATHENA/LIVE **
                                        If intQuestionFormatID <> 16 Then             'formatID that denotes a 5-field response
                                            intNoFields = 3             'yes/no/not applicable
                                        Else
                                            intNoFields = 5             'not at all/some of the time/most of the time/all of the time/not applicable
                                        End If
                                        If intNoChoices < intNoFields Then
                                            intNoChoices = intNoFields
                                            strHtml = strHtml & "<td width=""60"" class=""" & strStyle & """ style=""" & strStyle & """"
                                            strHtml = strHtml & ">Not<br>Applicable<br>&nbsp;</td>" & Chr(13)
                                        End If
                                        strHtml = strHtml & "</tr>" & Chr(13)
                                        Exit For
                                    End If
                                Next

                            Case 16          '  New Fill-in-the-blank type, ie, input box embedded inside question
                                'Create a new dataview instance on the table to iterate through question choices
                                Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)
                                intAccessKey = intAccessKey + 1
                                strAccessKey = GetAccessKey(intAccessKey)
                                strHtml = strHtml & "<tr>" & vbCrLf
                                strHtml = strHtml & "<td class=""questionText"" style=""questionText""><b>"
                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("iQuestionNumber") & ".&nbsp;&nbsp;</b>"
                                strQuestionText = QuestionChoiceDataView.Item(i).Row("txQuestion")
                                strQuestionChoice = "<input onClick=""setAnsweredSomething();"" type=text "
                                If QuestionChoiceDataView.Item(i).Row("iQuestionResponseTypeID") = 1 Then          'This was iFreeResposne ID
                                    intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                    If Not IsDBNull(intMax) Then          'its a question that needs specific min/max validation
                                        strQuestionChoice = strQuestionChoice & " size=3 maxlength=2 onBlur=""validateMinMax(this," & QuestionChoiceDataView.Item(i).Row("iMinimum") & ", " & intMax & ")"""
                                    Else
                                        strQuestionChoice = strQuestionChoice & " size=3 maxlength=2 onBlur=validate(this)"
                                    End If
                                Else
                                    strQuestionChoice = strQuestionChoice & " size=40 maxlength=500"
                                End If
                                strQuestionChoice = strQuestionChoice & " name=a" & intQuestionID
                                strQuestionChoice = strQuestionChoice & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & " value=""" & QuestionChoiceDataView.Item(i).Row("vchFreeResponse") & """>"
                                strQuestionChoice = strQuestionChoice & "<input type=hidden name=a" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">"
                                strQuestionText = Replace(strQuestionText, "<CHOICE>", strQuestionChoice, 1, -1, 1)
                                If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" And Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchFreeResponse")) Then
                                    answeredSomething = 1
                                End If
                                strHtml = strHtml & "(strQuestionText)"
                                strHtml = strHtml & "</td>" & vbCrLf
                                strHtml = strHtml & "</tr>" & vbCrLf
                            Case Else
                                strHtml = strHtml & "Sorry, error found on page!!"
                        End Select
                        strHtml = strHtml & "</td></tr>"
                    Else       'bAnserable <> true
                        i = i + 1       'still need to increment DATAVIEW counter if its not answerable
                        intDataViewIncrement = intDataViewIncrement + 1
                        'If it's the only question on the page we need to set answered something = 1
                        If QuestionTable.Rows.Count() = 1 Then
                            answeredSomething = 1
                            Session.Item("skipSave") = True
                            bCalQuestion = True
                        End If
                    End If       'end bAnswerable = true
                End If       'end bSubQuestion = 0
            End If    'end intHoldDataViewQuestionID <> intQuestionID
        Next    'QuestionRow In QuestionTable.Rows

        If (pageNumber > 0 Or blnCustomQuestion) And (Not bCalQuestion) Then    'the blnCustomQuestion is to ensure that the page doesnt blow out when printing just custom questions
            strHtml = strHtml & "</table></table>"
        Else
            strHtml = strHtml & "</table>"
        End If

        'print back and next links if not intro page
        'this will be pre-pended to the question string below
        If intPageStart <> 1 Then
            strLinks = "<table width=""100%"" cellpadding=0 cellspacing=0 border=0>"
            strLinks = strLinks & "<tr><td width=70>&nbsp;</td>"
            If iPageNumber > 1 Then
                strLinks = strLinks & "<td width=""15%"" align=right><a href=""javascript:goToPreviousPage();"" AccessKey=""A""><img height=15 width=99 src=/images/previous_page.gif border=0 alt=""Previous Page""></a></td>"
            Else    'page = 1
                strLinks = strLinks & "<td width=""15%"" align=right><img height=15 width=99 src=""/images/previous_page_grey.gif"" border=0 alt=""Previous Page""></td>"
            End If
            If intlastPage <> 0 Then
                strLinks = strLinks & "<td width=""15%""><a href=""javascript:goToNextPage();"" AccessKey=B><img height=15 width=77 src=""/images/next_page.gif"" border=0 alt=""Next Page""></a></td>"
            Else
                strLinks = strLinks & "<td width=""15%""><img height=15 width=77 src=""/images/next_page_grey.gif"" border=0 alt=""Next Page""></td>"
            End If
            strLinks = strLinks & "</table><br>"
        End If

        strHtml = strLinks & strHtml

        'Hidden field containing all Question IDs on this page.
        strHtml = strHtml & "<input type=hidden value=""" & strReqQuestions & """ name=""reqQuestions"">" & _
          "<input type=hidden value=""" & strQuestions & """ name=""ansQuestions"">" & _
          "<input type=hidden value=""" & strTextQuestion & """ name=""ansTextQuestions"">" & _
          "<input type=hidden value=""" & answeredSomething & """ name=""answeredFlag"">" & _
          "<input type=hidden value=""" & navFunction & """ name=""navFunction"">" & _
          "<input type=hidden value=""" & iPageNumber & """ name=""pageNumber"">" & _
          "<input type=hidden value=""" & blnSkipPattern1 & """ name=""blnSkipPattern1"">" & _
          "<input type=hidden value=""" & blnSkipPattern2 & """ name=""blnSkipPattern2"">" & _
          "<input type=hidden value=""" & blnSkipPattern3 & """ name=""blnSkipPattern3"">" & _
          "<input type=hidden value=""" & blnSkipPattern4 & """ name=""blnSkipPattern4"">" & _
          "<input type=hidden value=""" & blnSkipPattern5 & """ name=""blnSkipPattern5"">" & _
          "<input type=hidden value=""" & blnSkipPattern6 & """ name=""blnSkipPattern6"">" & _
            "<input type=hidden value=""" & blnSkipPattern7 & """ name=""blnSkipPattern7"">" & _
            "<input type=hidden value=""" & blnSkipPattern8 & """ name=""blnSkipPattern8"">" & _
            "<input type=hidden value=""" & blnSkipPattern9 & """ name=""blnSkipPattern9"">" & _
            "<input type=hidden value=""" & blnSkipPattern10 & """ name=""blnSkipPattern10"">" & _
            "<input type=hidden value=""" & blnSkipPattern11 & """ name=""blnSkipPattern11"">" & _
            "<input type=hidden value=""" & blnSkipPattern12 & """ name=""blnSkipPattern12"">" & _
            "<input type=hidden value=""" & blnSkipPattern13 & """ name=""blnSkipPattern13"">" & _
            "<input type=hidden value=""" & blnSkipPattern14 & """ name=""blnSkipPattern14"">"

        If intlastPage = 0 Then
            strHtml = strHtml & "<input type=hidden value=""1"" name=""EndSurvey"">"
        End If

        'TODO: Dispose of that datatable


        Return strHtml

    End Function

    Private Sub loadSurveyAnswers(ByVal iTestFormID As Integer)
        Dim dtAnswers As DataTable
        dtAnswers = getAnswers(iTestFormID)
        Session.Item("SurveyAnswers") = dtAnswers
    End Sub

    Sub saveSurvey()
        Dim responseAnswerArray As Object
        Dim strQuestionID As String
        Dim FreeResponse As Object
        Dim arrQuestionArray() As String
        Dim intQuestionID As String
        Dim intQuestionChoice As String
        Dim arrTextAnswerArray() As String
        Dim strTextQuestionID As String
        Dim k As Integer
        Dim i As Integer
        Dim strQuestionChoiceDropDown As String
        Dim intPos As Object
        Dim theAnswerArray() As String
        arrQuestionArray = Split(Request.Form.Item("ansQuestions"), ",")
        arrTextAnswerArray = Split(Request.Form.Item("ansTextQuestions"), ",")

        For i = 0 To UBound(arrQuestionArray)
            intQuestionID = arrQuestionArray(i)
            strQuestionID = "a" & intQuestionID

            If Not IsNothing(Request.Form("dropdown" & intQuestionID)) Then    'this is a dropdown question - capture questionchoice
                strQuestionChoiceDropDown = Request.Form("dropdown" & intQuestionID)
                Call saveAnswers(intQuestionID, strQuestionChoiceDropDown, Request.Form.Item(strQuestionID))
            Else    'not a drop down so go with the normal process of saving
                'Single Answer
                If Not IsNothing(Request.Form(strQuestionID)) Then
                    Call saveAnswers(intQuestionID, Request.Form.Item(strQuestionID), "")
                    strTextQuestionID = strQuestionID & "text" & Request.Form.Item(strQuestionID)
                    Call saveTextAnswers(intQuestionID, Request.Form.Item(strQuestionID), strTextQuestionID)
                    If intQuestionID = 6332 Then 'This is the gender question, save the gender to the session as SurveyFemale
                        'Set the survey gender here for modifying the text of two questions in Survey 1
                        setSurveyGender(Request.Form.Item(strQuestionID))
                    End If
                ElseIf Not IsNothing(Request.Form(strQuestionID & "_MC")) Then
                    'MC Multiple Response
                    theAnswerArray = Split(Request.Form.Item(strQuestionID & "_MC"), ",")
                    For k = 0 To UBound(theAnswerArray)
                        intQuestionChoice = Trim(theAnswerArray(k))
                        Call saveAnswers(intQuestionID, intQuestionChoice, "")
                        strTextQuestionID = strQuestionID & "text" & intQuestionChoice
                        Call saveTextAnswers(intQuestionID, intQuestionChoice, strTextQuestionID)

                        'Check to see if calendar question, if so, save those to session too...(--SEE BELOW)
                        'js 5/31/07: in testing surveys we found that if user was in same session from
                        ' survey 1 to survey 2 then the calendar responses which are saved to the
                        ' session below would prepopulate the calendar question in survey 2.
                        ' When this was first built to be used in the Drink Measure Survey, thats' 
                        ' exactly what we wanted...but we don't want that in the AEDu course surveys.
                        ' I will comment out the 'if...then' logic directly below and that should alleviate
                        ' the problem.

                        '3/2008 Use this feature to store the answers ONLY for use in writting out the highest drinking day on the next page.
                        If (Not IsNothing(Request.Form("calQuestion"))) And (Request.Form("calQuestion") = intQuestionID.ToString) Then
                            Call storeTextAnswers(strTextQuestionID)
                        End If
                    Next
                End If
            End If
        Next

        For i = 1 To UBound(arrTextAnswerArray)
            intQuestionID = arrTextAnswerArray(i)
            Call deleteAnswers(intQuestionID)
        Next
    End Sub
    Private Function makeArray(ByVal strArrayName As String)
        Dim strArray() As String

        If Trim(strArrayName) = "states" Then
            strArray = New String() {"Choose a State", "Alabama", "Alaska", "American Samoa", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "District Of Columbia", "Federated States of Micronesia", "Florida", "Georgia", "Guam", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Marshall Islands", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Northern Mariana Islands", "Ohio", "Oklahoma", "Oregon", "Palau", "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virgin Islands", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", "Did not graduate in the United States"}
        ElseIf Trim(strArrayName) = "gpa" Then
            strArray = New String() {"Choose a Grade", "A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-", "D+", "D", "D-", "F"}
        ElseIf Trim(strArrayName) = "time" Then
            strArray = New String() {"Choose a Time", "Never", "1 time", "2 times", "3 times", "4 times", "5 times", "6 times", "7 times", "8 times", "9 times", "10 times", "11 times", "12 times", "13 times", "14 times", "15 times", "16 times", "17 times", "18 times", "19 times", "20 or more times"}
        ElseIf Trim(strArrayName) = "drinks" Then
            strArray = New String() {"Choose number of drinks ", "1 drink", "2 drinks", "3 drinks", "4 drinks", "5 drinks", "6 drinks", "7 drinks", "8 drinks", "9 drinks", "10 drinks", "11 drinks", "12 drinks", "13 drinks", "14 drinks", "15 drinks", "16 drinks", "17 drinks", "18 drinks", "19 drinks", "20 or more drinks"}
        ElseIf Trim(strArrayName) = "recentday" Then
            strArray = New String() {"Choose a recent day", "Yesterday", "2 days ago", "3 days ago", "4 days ago", "5 days ago", "6 days ago", "7 days ago", "1-2 weeks ago", "3-4 weeks ago", "1-2 months ago", "3-4 months ago", "5-6 months ago", "7-9 months ago", "10-12 months ago"}
        ElseIf Trim(strArrayName) = "consume" Then
            strArray = New String() {"Choose drinks consumed ", "I don't experience this when I drink", "1 drink", "2 drinks", "3 drinks", "4 drinks", "5 drinks", "6 drinks", "7 drinks", "8 drinks", "9 drinks", "10 or more drink"}
        ElseIf Trim(strArrayName) = "hours" Then
            strArray = New String() {"Choose a Time of Day", "6:00 a.m.", "6:30 a.m.", "7:00 a.m.", "7:30 a.m.", "8:00 a.m.", "8:30 a.m.", "9:00 a.m.", "9:30 a.m.", "10:00 a.m.", "10:30 a.m.", "11:00 a.m.", "11:30 a.m.", "12:00 p.m.", "12:30 p.m.", "1:00 p.m.", " 1:30 p.m.", "2:00 p.m.", "2:30 p.m.", " 3:00 p.m.", "3:30 p.m.", "4:00 p.m.", "4:30 p.m.", " 5:00 p.m.", "5:30 p.m.", "6:00 p.m.", "6:30 p.m.", "7:00 p.m.", "7:30 p.m.", "8:00 p.m.", "8:30 p.m.", "9:00 p.m.", "9:30 p.m.", "10:00 p.m.", "10:30 p.m.", "11:00 p.m.", "11:30 p.m.", "12:00 a.m.", "12:30 a.m.", "1:00 a.m.", "1:30 a.m.", "2:00 a.m.", "2:30 a.m.", "3:00 a.m.", "3:30 a.m.", "4:00 a.m.", "4:30 a.m.", "5:00 a.m.", "5:30 a.m.", "Don''t know"}
        ElseIf Trim(strArrayName) = "location" Then
            strArray = New String() {"Choose a Location", "Your room", "College Residence hall (not your room)", "Fraternity/sorority house (not your room)", "Off-campus house or apartment (not your room)", "Parent/guardian’s home (not your room)", "Another family member's home (not parent/guardian)", "On-campus pub", "Campus academic building", "On-campus dance or concert", "Tailgating party", "On-campus sporting event", "On-campus outdoor setting", "Off-campus bar, tavern, club, or restaurant", "Off-campus outdoor setting", "Automobile", "Substance-free residence hall (not your room)", "Outdoors (e.g., at the beach, while camping)", "Other location"}
        ElseIf Trim(strArrayName) = "days" Then
            strArray = New String() {"Choose number of days", "0 days", "1 day", "2 days", "3 days", "4 days", "5 days", "6 days", "7 days", "8 days", "9 days", "10 days", "11 days", "12 days", "13 days", "14 days"}
        ElseIf Trim(strArrayName) = "error" Then
            strArray = New String() {"Error retrieving values"}
        End If

        'return array
        makeArray = strArray

    End Function
    Private Sub deleteAnswers(ByVal intQuestionID As Integer)
        'Run the DataAccess class and delete text Survey answers
        Dim objDeleteAnswers As New DataAccess.DAAssessment
        objDeleteAnswers.DeleteAnswers(intQuestionID, iAssessmentID)
    End Sub
    Sub saveTextAnswers(ByVal intQuestionID As Integer, ByVal intQuestionChoiceID As Integer, ByVal strTextQuestionID As String)
        'Dim cleanSQL() As String
        Dim strTextAnswer As String
        If Not IsNothing(Request.Form.Item(strTextQuestionID)) Then
            strTextAnswer = Trim(Request.Form.Item(strTextQuestionID))
            If (Not IsNothing(strTextAnswer)) Then    '1/3/01
                strTextAnswer = cleanSQL(strTextAnswer)
            End If
            Call saveAnswers(intQuestionID, intQuestionChoiceID, strTextAnswer)

        End If
    End Sub
    Function cleanSQL(ByRef text As String) As String
        If text <> "" Then
            text = Replace(text, "'", "''")
            text = Replace(text, "?", "")
            cleanSQL = text
        Else
            cleanSQL = ""
        End If
    End Function
    Function getSubQuestion(ByVal subQuestionID As Integer, ByVal QuestionTable As DataTable) As String
        Dim QuestionColumn As DataColumn
        Dim QuestionRow As DataRow
        For Each QuestionRow In QuestionTable.Rows
            If QuestionRow.Item("iQuestionID") = subQuestionID Then
                strSubQuestions = strSubQuestions & "<input type=text value="""
                strSubQuestions = strSubQuestions & QuestionRow.Item("vchFreeResponse")
                strSubQuestions = strSubQuestions & """ name=a" & subQuestionID & "text" & QuestionRow.Item("iQuestionChoiceID") & " size=20 maxlength=500>"
                strSubQuestions = strSubQuestions & "<input type=hidden name=a" & subQuestionID & " value=" & QuestionRow.Item("iQuestionChoiceID") & ">"
                strQuestions = strQuestions & subQuestionID & ","
                strTextQuestion = strTextQuestion & "," & subQuestionID
            End If
        Next
        Return strSubQuestions

    End Function

    Function getSubQuestion2(ByVal subQuestionID) As String

        Dim DRSubQuestion As DataSet
        Dim QuestionTable As DataTable
        Dim QuestionColumn As DataColumn
        Dim QuestionRow As DataRow
        Dim objSubQuestion As New Content.PopulateLists
        'this just gets the same datareader for comparison to find the subquestions
        DRSubQuestion = objSubQuestion.getQuestions(iSegmentID, iAssessmentID, pageNumber)

        If DRSubQuestion.Tables.Count > 0 Then    'dataset contains a table...proceed
            For Each QuestionTable In DRSubQuestion.Tables
                For Each QuestionRow In QuestionTable.Rows
                    If QuestionRow.Item("iQuestionID") = subQuestionID Then
                        strSubQuestions = strSubQuestions & "<input type=text value="""
                        strSubQuestions = strSubQuestions & QuestionRow.Item("vchFreeResponse")
                        strSubQuestions = strSubQuestions & """ name=a" & subQuestionID & "text" & QuestionRow.Item("iQuestionChoiceID") & " size=20 maxlength=500>"
                        strSubQuestions = strSubQuestions & "<input type=hidden name=a" & subQuestionID & " value=" & QuestionRow.Item("iQuestionChoiceID") & ">"
                        strQuestions = strQuestions & subQuestionID & ","
                        strTextQuestion = strTextQuestion & "," & subQuestionID
                    End If
                Next
            Next
        End If

        Return strSubQuestions

    End Function
    Private Sub saveAnswers(ByVal intQuestionID As Integer, ByVal iQuestionChoiceID As Integer, ByVal strTextAnswer As String)
        'Run the DataAccess class and save Survey answers
        Dim objSaveAnswers As New DataAccess.DAAssessment
        objSaveAnswers.SaveAnswers(intQuestionID, iAssessmentID, iQuestionChoiceID, strTextAnswer)
    End Sub
    Private Sub setUserClassification()
        'Run the DataUser class and set user classification based on Survey answers
        Dim objSetUserClassification As New DataAccess.DAUsers
        objSetUserClassification.SetUserClassification(iUserAccountID, iCourseID)
    End Sub
    Private Sub socClassify(ByVal iUserAccountID As Integer)
        'Run the DataUser class and set SOC classification based on Survey 3 answers
        Dim objSetUserClassification As New DataAccess.DAUpdateUserData
        objSetUserClassification.classifySOC(iUserAccountID)
    End Sub
    Private Sub setSurveyDate()
        'Run the DataUser class and set user classification based on Survey answers
        Dim objSetUserClassification As New DataAccess.DAUsers
        objSetUserClassification.SetSurveyDate(iUserAccountID, iSegmentID, iAssessmentID)
    End Sub
    Function GetAccessKey(ByVal intAccessKey As Integer) As String

        Select Case intAccessKey
            Case 1
                strAccessKey = "C"
            Case 2
                strAccessKey = "D"
            Case 3
                strAccessKey = "E"
            Case 4
                strAccessKey = "F"
            Case 5
                strAccessKey = "G"
            Case 6
                strAccessKey = "H"
            Case 7
                strAccessKey = "I"
            Case 8
                strAccessKey = "J"
            Case 9
                strAccessKey = "K"
            Case 10
                strAccessKey = "L"
            Case 11
                strAccessKey = "M"
            Case 12
                strAccessKey = "N"
            Case 13
                strAccessKey = "O"
            Case 14
                strAccessKey = "P"
            Case 15
                strAccessKey = "Q"
            Case 16
                strAccessKey = "R"
            Case 17
                strAccessKey = "S"
            Case 18
                strAccessKey = "T"
            Case 19
                strAccessKey = "U"
            Case 20
                strAccessKey = "V"
            Case 21
                strAccessKey = "W"
            Case 22
                strAccessKey = "X"
            Case 23
                strAccessKey = "Y"
            Case 24
                strAccessKey = "Z"
            Case 25
                strAccessKey = ","
        End Select

        GetAccessKey = strAccessKey

    End Function
    Function isValueDefined(ByRef strFieldName As String) As Boolean
        Dim strMCFieldName As String
        strMCFieldName = strFieldName & "_MC"
        If Trim(Request.Form.Item(strFieldName)) = "" And Trim(Request.Form.Item(strMCFieldName)) = "" Then
            isValueDefined = False
        Else
            isValueDefined = True
        End If

    End Function
    Function GetTestTypeID(ByVal intSegmentID) As Integer
        Dim intTestType As Integer
        Dim intSurveyType As Integer
        Dim DRTestType As SqlClient.SqlDataReader
        'get the list values
        Dim objDRTestType As New DataAccess.DANavigation
        intTestType = objDRTestType.getTestType(intSegmentID)

        GetTestTypeID = intTestType
    End Function
    'This subroutine will run the procedure that copies the Demographic answers from the Interview to the PreSurvey
    Private Sub runPrePopulate()
        'Run the DataUser class and prepopulate users assessment with interview responses
        Dim objPrePopulate As New DataAccess.DAAssessment
        objPrePopulate.PrePopulate(iUserAccountID)
    End Sub

    Private Function isTestFormDone() As Boolean
        'Check to see if the current assessment is complete
        Dim objAssessmentDone As New DataAccess.DAAssessment
        Return objAssessmentDone.checkAssessmentDone(iAssessmentID)
    End Function
    Private Sub checkSurvey()
        Dim bDone As Boolean
        bDone = isTestFormDone()
        If Not bDone Then
            nullifyAnswers()
        Else
            moveNext()
        End If
    End Sub

    Private Sub nullifyAnswers()
        'Run the DataUser class and nullify previous answers if user didn't complete survey
        Dim objNullifyAnswers As New DataAccess.DAUsers
        objNullifyAnswers.NullifyAnswers(iAssessmentID)
    End Sub
    Private Function checkFinal() As Integer
        Dim intFinalAnswers
        'Run the DataUser class and check if user completed final before starting post survey
        Dim objcheckFinal As New DataAccess.DAAssessment
        intFinalAnswers = objcheckFinal.CheckFinal(iUserAccountID)

        checkFinal = intFinalAnswers
    End Function
    Public Sub readSession(ByRef iUserAccountID As Integer, ByRef iMediaPlayerID As Integer, ByRef iConnectionSpeedID As Integer, ByRef iAssessmentID As Integer, ByRef iSegmentID As Integer, ByRef strSegmentName As String, ByRef iClientID As Integer, ByRef iTestID As Integer, ByRef iScreenHeight As Integer, ByRef bJava As Boolean, ByRef strNavFunction As String, ByRef iCourseID As Integer, ByRef iClassificationID As Integer, ByRef iCourseTypeID As Integer, ByRef iClientCourseID As Integer)
        Dim objSession As StateMgt.Session = New StateMgt.Session
        objSession.readSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
    End Sub

    Public Sub updateSession(ByVal iUserAccountID As Integer, ByVal iMediaPlayerID As Integer, ByVal iConnectionSpeedID As Integer, ByVal iAssessmentID As Integer, ByVal iSegmentID As Integer, ByVal strSegmentName As String, ByVal iClientID As Integer, ByVal iTestID As Integer, ByVal iScreenHeight As Integer, ByVal bJava As Boolean, ByVal strNavFunction As String, ByVal iCourseID As Integer, ByVal iClassificationID As Integer, ByVal iCourseTypeID As Integer, ByVal iClientCourseID As Integer)
        Dim objSession = New StateMgt.Session
        objSession.updateSession(iUserAccountID, iMediaPlayerID, iConnectionSpeedID, iAssessmentID, iSegmentID, strSegmentName, iClientID, iTestID, iScreenHeight, bJava, strNavFunction, iCourseID, iClassificationID, iCourseTypeID, iClientCourseID)
    End Sub
    Public Sub updateSessionValue(ByVal strIndex As String, ByVal strValue As String)
        Dim objSession = New StateMgt.Session
        objSession.updateSessionValue(strIndex, strValue)
    End Sub

    Private Sub addItemToContext(ByVal vchPageNumber As String, ByVal arrAnswers() As String)
        Context.Items.Add("vchPageNumber", arrAnswers)
    End Sub

    Private Function readContext() As Array
        Return Context.Items.Item("Answers")
    End Function
    Public Function skipQsetEnrgDrinks2()
        'for skip pattern 14: 'Qset Energy Drinks question #2
        Dim blnSkip As Boolean
        If intTestTypeID = 19 Or intTestTypeID = 1 Then 'survey 1
            If pageNumber = 2 Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 3 'to last page - pg 5
            End If

        End If

        If blnSkip Then
            blnSkipPattern14 = True
        End If
        skipQsetEnrgDrinks2 = pageNumber
    End Function
    Public Function skipQsetAlSettings1()
        'for skip pattern 11: 'Qset Alcohol Settings question #1
        Dim blnSkip As Boolean
        If intTestTypeID = 19 Or intTestTypeID = 1 Then 'survey 1
            If pageNumber = 2 Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 8 'to last page - pg 9
            End If
        End If

        If blnSkip Then
            blnSkipPattern11 = True
        End If
        skipQsetAlSettings1 = pageNumber
    End Function
    Public Function skipQsetAlSettings5()
        'for skip pattern 12: ''Qset Alcohol Settings question #5
        Dim blnSkip As Boolean
        If intTestTypeID = 19 Or intTestTypeID = 1 Then 'survey 1
            If pageNumber = 4 Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 2 'to last page - pg 6
            End If
        End If

        If blnSkip Then
            blnSkipPattern12 = True
        End If
        skipQsetAlSettings5 = pageNumber
    End Function
    Public Function skipQsetAlSettings7()
        'for skip pattern 13: 'Qset Alcohol Settings question #7
        Dim blnSkip As Boolean
        If intTestTypeID = 19 Or intTestTypeID = 1 Then 'survey 1
            If pageNumber = 6 Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 2 'to  page - pg 7
            End If
        End If

        If blnSkip Then
            blnSkipPattern13 = True
        End If
        Return pageNumber
    End Function
    Public Function skipNoDrinksPastYear()
        'for skip pattern 1: 'During the past year...have you consumed any alcohol..."
        Dim blnSkip As Boolean
        If intTestTypeID = 1 Then 'survey 1
            If pageNumber = 2 And (iCourseTypeID = 1 Or iCourseTypeID = 12 Or iCourseTypeID = 3) Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 6 'QID 6303 on page 8
            ElseIf pageNumber = 2 And (iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Or iCourseTypeID = 4) Then 'preMat/Summer/SAEdu Summer
                blnSkip = True
                pageNumber = pageNumber + 6 'QID 6303 on page 8
                'ElseIf pageNumber = 2 And (iCourseTypeID = 4) Then 'preMat CONTROL
                '    blnSkip = True
                '    pageNumber = pageNumber + 5 'to Q#10/pg 7
            ElseIf pageNumber = 7 And iCourseTypeID = 8 Then 'sanctions
                blnSkip = True
                pageNumber = pageNumber + 6
            End If
        ElseIf intTestTypeID = 10 Then 'survey 2
            If pageNumber = 8 And (iCourseTypeID = 1 Or iCourseTypeID = 8 Or iCourseTypeID = 12) Then  'postMat or Sanctions or SAEdu postMat
                blnSkip = True
                pageNumber = pageNumber + 2
            ElseIf pageNumber = 6 And (iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'preMat/summer/saedu summer
                blnSkip = True
                pageNumber = pageNumber + 2
            End If
        ElseIf intTestTypeID = 11 Then 'survey 3
            If pageNumber = 3 Then
                blnSkip = True
                pageNumber = pageNumber + 5 'to Q#16
                'ElseIf pageNumber = 2 And (iCourseTypeID = 4) Then 'preMat CONTROL
                '    blnSkip = True
                '    pageNumber = pageNumber + 5 'to Q#10/pg 7
            End If
        End If

        If blnSkip Then
            blnSkipPattern1 = True
        End If
        Return pageNumber
    End Function
    Public Function skipNoDrinksPastTwoWeeks()
        'for skip pattern 2: 'During the past two weeks...have you consumed any alcohol..."
        Dim blnSkip As Boolean
        If intTestTypeID = 1 Then 'survey 1
            If pageNumber = 3 And (iCourseTypeID = 1 Or iCourseTypeID = 12 Or iCourseTypeID = 3) Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 3 'to Q#15/pg 8
            ElseIf pageNumber = 3 And (iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17 Or iCourseTypeID = 4) Then 'preMat/Summer/SAEdu Summer
                blnSkip = True
                pageNumber = pageNumber + 3 'to Q#15/pg 8
            ElseIf pageNumber = 7 And iCourseTypeID = 8 Then 'sanctions
                blnSkip = True
                pageNumber = pageNumber + 6
                'ElseIf pageNumber = 3 And iCourseTypeID = 4 Then 'PreMat Control
                '    blnSkip = True
                '    pageNumber = pageNumber + 3 'Q#9/Pg 6
            End If
        ElseIf intTestTypeID = 10 Then 'survey 2
            If pageNumber = 8 And (iCourseTypeID = 1 Or iCourseTypeID = 8 Or iCourseTypeID = 12) Then  'postMat or Sanctions or SAEdu postMat
                blnSkip = True
                pageNumber = pageNumber + 2
            ElseIf pageNumber = 6 And (iCourseTypeID = 2 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'preMat/Sumer/Summer SAEdu
                blnSkip = True
                pageNumber = pageNumber + 2
            End If
        ElseIf intTestTypeID = 11 Then 'survey 3
            If pageNumber = 4 Then
                blnSkip = True
                pageNumber = pageNumber + 3 'to Q#14/pg 7
            ElseIf pageNumber = 3 And iCourseTypeID = 4 Then 'PreMat Control
                blnSkip = True
                pageNumber = pageNumber + 3 'Q#9/Pg 6
            End If
        End If

        If blnSkip Then
            blnSkipPattern2 = True
        End If
        Return pageNumber
    End Function
    Public Function skipUSCitizen()
        'for skip pattern 3: 'Are you a US citizen"
        Dim blnSkip As Boolean
        If intTestTypeID = 1 Then 'survey 1
            If pageNumber = 10 And (iCourseTypeID = 1 Or iCourseTypeID = 3) Then 'postMat types
                blnSkip = True
                pageNumber = pageNumber + 2 'to Q#26?
            ElseIf pageNumber = 10 And (iCourseTypeID = 2 Or iCourseTypeID = 16) Then 'preMat/Summer/SAEdu Summer
                blnSkip = True
                pageNumber = pageNumber + 2 'to Q#25?
            ElseIf pageNumber = 11 And (iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 17) Then 'All SAEDU types
                blnSkip = True
                pageNumber = pageNumber + 2 'to Q#26?
            ElseIf pageNumber = 7 And iCourseTypeID = 8 Then 'sanctions
                blnSkip = True
                pageNumber = pageNumber + 6
            End If
        End If

        If blnSkip Then
            blnSkipPattern3 = True
        End If
        Return pageNumber
    End Function
    Public Function skipQsetDrinkExp1()
        'for skip pattern 6: 'Qset Drinking Exp. can follow survey 1 and survey 3
        Dim blnSkip As Boolean
        If pageNumber = 2 And (iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'postMat types Then
            blnSkip = True
            pageNumber = pageNumber + 7 'to last page - pg 9
        End If

        If blnSkip Then
            blnSkipPattern6 = True
        End If
        skipQsetDrinkExp1 = pageNumber
    End Function
    Public Function skipQsetDiary1()
        'for skip pattern 7: 'Qset Diary. can follow survey 1 and survey 3
        Dim blnSkip As Boolean
        If pageNumber = 2 And (iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'postMat types Then
            blnSkip = True
            pageNumber = pageNumber + 8 'from page 2 to last page - pg 10
        End If

        If blnSkip Then
            blnSkipPattern7 = True
        End If
        skipQsetDiary1 = pageNumber
    End Function
    Public Function skipQsetDiary8()
        'for skip pattern 8: 'Qset Diary. can follow survey 1 and survey 3
        Dim blnSkip As Boolean
        If pageNumber = 4 And (iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'postMat types Then
            blnSkip = True
            pageNumber = pageNumber + 4  'from page 4/question 8 to pg 8/question 23
        End If

        If blnSkip Then
            blnSkipPattern8 = True
        End If
        skipQsetDiary8 = pageNumber
    End Function
    Public Function skipQsetDiary13()
        'for skip pattern 9: 'Qset Diary. can follow survey 1 and survey 3
        Dim blnSkip As Boolean
        If pageNumber = 5 And (iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'postMat types Then
            blnSkip = True
            pageNumber = pageNumber + 3  'from page 5/question 13 to to pg 8/question 23
        End If

        If blnSkip Then
            blnSkipPattern9 = True
        End If
        skipQsetDiary13 = pageNumber
    End Function

    Public Function skipQsetDiary18()
        'for skip pattern 10: 'Qset Diary. can follow survey 1 and survey 3
        Dim blnSkip As Boolean
        If pageNumber = 6 And (iCourseTypeID = 1 Or iCourseTypeID = 2 Or iCourseTypeID = 12 Or iCourseTypeID = 13 Or iCourseTypeID = 16 Or iCourseTypeID = 17) Then 'postMat types Then
            blnSkip = True
            pageNumber = pageNumber + 2 'from page 6/question 18 to pg 8/question 23
        End If

        If blnSkip Then
            blnSkipPattern10 = True
        End If
        skipQsetDiary18 = pageNumber
    End Function
    'for skip pattern 4 ON SURVEY 2: 'Did the user answer 'no' to Q#5/Survey 1...if so, on survey 2 skip Q#5 and 6 and go to Q#7 "
    Public Function skipSurvey2NoDrinksPastYear()
        Dim intSurvey2NoDrinksPastYear = 25033  'THE 'NO' CHOICE TO Q#5/SURVEY 1  questionID = 6211 on both ATHENA & LIVE !!!  
        Dim iTestFormID As Integer
        Dim dr As DataRow
        Dim dt As DataTable
        Dim iCount As Integer
        Dim blnSkip As Boolean

        iTestFormID = getTestFormID()
        If iTestFormID = 8 Then 'we are on the postmat so get the postmat survey 1 responses
            iTestFormID = 6 'postmat survey 1
        ElseIf iTestFormID = 111 Then 'we are on the premat so get the premat survey 1 responses
            iTestFormID = 112 'premat survey 1
        End If

        dt = getAnswers(iTestFormID)
        iCount = dt.Rows.Count()

        If Not IsNothing(dt) Then
            iCount = dt.Rows.Count()
        Else
            iCount = 0
        End If

        If iCount > 0 Then
            dr = dt.Rows.Item(i)
            Do While iCount > i  'the response we are looking for comes from survey 1/question #5 pre/post
                dr = dt.Rows.Item(i)
                If dr.Item("iQuestionChoiceID") = intSurvey2NoDrinksPastYear Then 'user answered 'no' to Q#5 on Survey 1
                    blnSkip = True
                    pageNumber = pageNumber + 2 'skip to Q#7/PG 5 on survey 2
                    Exit Do
                End If
                i = i + 1
                If iCount > i Then
                    dr = dt.Rows.Item(i)
                End If
            Loop
        Else
            blnSkip = False
        End If

        If blnSkip Then
            blnSkipPattern4 = True
        Else
            blnSkipPattern4 = False
        End If

        skipSurvey2NoDrinksPastYear = pageNumber
    End Function
    Public Function getCourseTypeID()
        Dim intCType As Integer
        'get coursetypeid to be used on page
        Dim objCType As New DataAccess.DAClient
        intCType = objCType.getCourseTypeID(iCourseID, iClientID)

        updateSessionValue("iCourseType", intCType)

        getCourseTypeID = intCType
    End Function

    Private Function checkForAnswer(ByVal iQuestionChoiceID As Integer, ByRef vchFreeResponse As String, ByRef dvAnswers As DataView) As Boolean
        Dim intRowIndex As Integer
        Dim strTestAnswer As String
        intRowIndex = dvAnswers.Find(iQuestionChoiceID)
        If intRowIndex = -1 Then
            vchFreeResponse = ""
            Return False
        Else
            vchFreeResponse = dvAnswers.Item(intRowIndex).Row("vchFreeResponse")
            vchDropDown = vchFreeResponse
            Return True
        End If

    End Function
    Private Function getXMLDT(ByVal iTestFormID As Integer, ByVal iPageNumber As Integer) As DataTable
        Dim objSurvey As Content.Survey = New Content.Survey
        Return objSurvey.readSurveyXML(iTestFormID, iPageNumber, iErrorCode)
    End Function
    Private Sub createXMLSurvey(ByVal iTestFormID As Integer, ByVal iPageNumber As Integer)
        Dim objSurvey As Content.Survey = New Content.Survey
        objSurvey.createSurveyXML(iTestFormID, iPageNumber)
    End Sub
    Private Function getTestFormID() As Integer
        Dim objTestForm As DataAccess.DAAssessment = New DataAccess.DAAssessment
        Return objTestForm.getTestForm(iSegmentID)
    End Function
    Private Function getAnswers(ByVal iTestFormID As Integer) As DataTable
        Dim objAnswers As DataAccess.DAAssessment = New DataAccess.DAAssessment
        Return objAnswers.getAssessmentResponses(iUserAccountID, iTestFormID)
    End Function
    Private Sub addAnswers(ByRef answerTable As Hashtable, ByVal strQID As String, ByVal strQCID As String, ByVal strVchChoice As String)
        Dim strKey As String
        strKey = strQID & "_" & strQCID
        answerTable.Add(strKey, strVchChoice)
    End Sub
    Private Sub startTimer(ByRef startTime As TimeSpan)
        startTime = Now.TimeOfDay
    End Sub

    '-- Start of Functions for the Calendar Question Type


    Function createCalendarGrid(ByVal QuestionChoiceDataView As DataView, ByVal AnswerView As DataView, ByVal iQuestionID As Integer, ByRef intDataViewIncrement As Integer, ByRef intHoldDataViewQuestionID As Integer, ByRef strReqQuestions As String, ByRef strTextQuestion As String, ByRef strHiddenFieldValues As String) As String
        Dim strOutput As String
        Dim dtToday As Date = Today()
        'Dim dtToday As Date = DateSerial(2007, 1, 10)
        Dim dtstart As Date
        Dim dtEnd As Date
        Dim dtLastDayofMonth As Date
        Dim dtCurrentDate As Date
        Dim iMonth As Integer
        Dim iDate As Integer
        Dim iDateEnd As Integer
        Dim strMonth As String
        Dim iQuestionChoiceID As Integer
        Dim iDay As Integer
        Dim iDayOfWeek As Integer
        Dim iCount As Integer = 1
        Dim bTwoMonths As Boolean = False
        Dim i As Integer = 0
        Dim intDataViewQuestionID As Integer
        Dim FreeResponse As String
        Dim responseSelected As Boolean
        Dim intMax As Integer
        Dim iMinimum As Integer
        Dim strHiddenFields As String

        If Request("navFunction") = "back" Then
            'js 4/16/08 -- to address problem with user going back to calendar and changing values...if
            ' they change the 'high' value upon going back, then when they move forward again and see
            ' question #10 they will see thier 'original' high value, not the 'new' high value they just input
            'So, maybe if we remove/clear the session values...and 0-out the calCounter so there should never
            ' be more than 14 calCounter values...
            RemoveCalSession()
            Session.Item("calCounter") = 0
        End If

        dtstart = DateAdd(DateInterval.Day, -14, dtToday)    'Subtract 2 weeks days from today to find starting date for question
        dtEnd = DateAdd(DateInterval.Day, 13, dtstart)
        iMonth = Month(dtEnd)

        If iMonth <> Month(dtstart) Then
            'This crosses two months, need to create 2 grids.
            iMonth = Month(dtstart)
            bTwoMonths = True
        End If

        strMonth = MonthName(iMonth)

        iDateEnd = DatePart(DateInterval.Day, dtEnd)
        dtCurrentDate = DateSerial(DatePart(DateInterval.Year, dtstart), iMonth, 1)    'Start on the First day of the month containing the dtStart
        If iMonth = 12 Then
            dtLastDayofMonth = DateSerial(DatePart(DateInterval.Year, dtstart), 12, 31)
        Else
            dtLastDayofMonth = DateAdd(DateInterval.Day, -1, (DateSerial(DatePart(DateInterval.Year, dtstart), iMonth + 1, 1)))
        End If

        strOutput = startTable(strMonth)

        'Create an array to store both the date and corresponding input field name
        Dim arrDates(14, 1) As String

        Do While dtCurrentDate <= dtLastDayofMonth
            iDayOfWeek = DatePart(DateInterval.Weekday, dtCurrentDate)
            iDate = DatePart(DateInterval.Day, dtCurrentDate)
            If iDate = 1 And iDayOfWeek <> 1 Then
                startRow(strOutput)
                Do While iCount < iDayOfWeek
                    addStaticField(strOutput, 0)
                    iCount = iCount + 1
                Loop
            Else
                If iDayOfWeek = 1 Then startRow(strOutput)
            End If

            If dtCurrentDate >= dtstart And dtCurrentDate <= dtEnd Then
                'Create active cell
                'Get the QustionChoiceID, iMinimum and intMax for this choice from the dv
                'Create an array of dates/questionchoices and save in session.
                i = intDataViewIncrement
                If i < QuestionChoiceDataView.Count() Then
                    intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                    If intDataViewQuestionID = iQuestionID Then
                        intDataViewIncrement = intDataViewIncrement + 1
                        intHoldDataViewQuestionID = intDataViewQuestionID
                        'Now check for this answer in the dvAnswers DataView
                        responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, AnswerView)
                        If responseSelected Then answeredSomething = 1
                        intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                        iMinimum = QuestionChoiceDataView.Item(i).Row("iMinimum")
                        If strReqQuestions <> "" Then
                            strReqQuestions = strReqQuestions & ",a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                        Else
                            strReqQuestions = "a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                        End If
                        If strHiddenFieldValues <> "" Then
                            strHiddenFieldValues = strHiddenFieldValues & "," & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                        Else
                            strHiddenFieldValues = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                        End If
                        strTextQuestion = strTextQuestion & "," & iQuestionID
                        addActiveField(strOutput, iQuestionID, QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), iDate, iMinimum, intMax, responseSelected, FreeResponse)
                        arrDates(i, 0) = dtCurrentDate.ToString
                        arrDates(i, 1) = "a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                    End If
                End If
            Else
                'Create static cell
                addStaticField(strOutput, iDate)
            End If

            If iDayOfWeek = 7 Then endRow(strOutput)
            dtCurrentDate = DateAdd(DateInterval.Day, 1, dtCurrentDate)
        Loop
        If bTwoMonths Then
            'Show the second Month
            If iDayOfWeek = 99 Then
                'Start a whole new calendar
            Else
                'pick up with fields where left off??
                iMonth = Month(dtEnd)
                Do While dtCurrentDate <= dtEnd
                    iDayOfWeek = DatePart(DateInterval.Weekday, dtCurrentDate)
                    iDate = DatePart(DateInterval.Day, dtCurrentDate)
                    If iDayOfWeek = 1 Then startRow(strOutput)


                    If dtCurrentDate >= dtstart And dtCurrentDate <= dtEnd Then
                        'Create active cell
                        i = intDataViewIncrement
                        If i < QuestionChoiceDataView.Count() Then
                            intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                            If intDataViewQuestionID = iQuestionID Then
                                intDataViewIncrement = intDataViewIncrement + 1
                                intHoldDataViewQuestionID = intDataViewQuestionID
                                'Now check for this answer in the dvAnswers DataView
                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, AnswerView)
                                intMax = QuestionChoiceDataView.Item(i).Row("iMaximum")
                                iMinimum = QuestionChoiceDataView.Item(i).Row("iMinimum")
                                If strReqQuestions <> "" Then
                                    strReqQuestions = strReqQuestions & ",a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                Else
                                    strReqQuestions = "a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                End If
                                If strHiddenFieldValues <> "" Then
                                    strHiddenFieldValues = strHiddenFieldValues & "," & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                Else
                                    strHiddenFieldValues = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                End If
                                strTextQuestion = strTextQuestion & "," & iQuestionID
                                addActiveField(strOutput, iQuestionID, QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), iDate, iMinimum, intMax, responseSelected, FreeResponse)
                                arrDates(i, 0) = dtCurrentDate.ToString
                                arrDates(i, 1) = "a" & iQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")

                            End If
                        End If
                    Else
                        'Create static cell
                        addStaticField(strOutput, iDate)
                    End If

                    If iDayOfWeek = 7 Then endRow(strOutput)
                    dtCurrentDate = DateAdd(DateInterval.Day, 1, dtCurrentDate)
                Loop
            End If

            If iDayOfWeek < 7 Then
                iCount = iDayOfWeek
                If bTwoMonths Then
                    iDate = DatePart(DateInterval.Day, dtCurrentDate)
                Else
                    iDate = 1
                End If
                Do While iCount < 7
                    addStaticField(strOutput, iDate)
                    iCount = iCount + 1
                    iDate = iDate + 1
                Loop
                endRow(strOutput)
            End If

        End If

        endTable(strOutput)


        strOutput = strOutput & strHiddenFields


        Session.Item("DateArray") = arrDates


        Return strOutput

    End Function


    Private Sub endTimer(ByVal startTime As TimeSpan)
        Dim ElapsedTime As TimeSpan
        ElapsedTime = Now.TimeOfDay.Subtract(startTime)
        displayTimer(ElapsedTime)
    End Sub

    Private Sub displayTimer(ByVal ElapsedTime As TimeSpan)
        Trace.Write("Timer", "Run Time: " & ElapsedTime.ToString)
    End Sub

    Private Function startTable(ByVal strMonth As String) As String
        Dim strTable As String
        strTable = "<table cellspacing=0 cellpadding=1 border=1 width=""500"" align=""center""><tr><td align=""center"" colspan=""7"">" & strMonth & "</td></tr><tr>" & vbCrLf _
        & "<TR><TD align=""center"">Sunday</TD><TD align=""center"">Monday</TD><TD align=""center"">Tuesday</TD><TD align=""center"">Wednesday</TD><TD align=""center"">Thursday</TD><TD align=""center"">Friday</TD><TD align=""center"">Saturday</TD></TR>"

        Return strTable
    End Function

    Private Sub endTable(ByRef strTable As String)
        strTable = strTable & "</table>"
    End Sub

    Private Sub startRow(ByRef strTable As String)
        strTable = strTable & vbCrLf & "<TR>"
    End Sub

    Private Sub endRow(ByRef strTable As String)
        strTable = strTable & "</tr>" & vbCrLf
    End Sub

    Private Sub addActiveField(ByRef strTable As String, ByVal iQuestionID As Integer, ByVal iQuestionChoiceID As Integer, ByVal iDate As Integer, ByVal iMinimum As Integer, ByVal intMax As Integer, ByVal ResponseSelected As Boolean, ByVal FreeResponse As String)
        Dim strCell As String
        Dim iAnswerCount As Integer = 0
        Dim name As String = "a" & iQuestionID.ToString & "text" & iQuestionChoiceID.ToString
        Dim strAnswer As String = ""
        If ResponseSelected Then
            strAnswer = FreeResponse
        Else
            'If there are answers in the session, use those. This should pre=populate the second question, but only if not already answered.
            If Not IsNothing(Session.Item("calCounter")) Then
                If Not IsNothing(Session.Item("ansCounter")) Then
                    'If Session.Item("calCounter") = 14 Then
                    iAnswerCount = Session.Item("ansCounter")
                Else
                    iAnswerCount = 1
                End If
                If iAnswerCount < 14 Then
                    Session.Item("ansCounter") = iAnswerCount + 1
                Else
                    Session.Item("ansCounter") = 1
                End If
                answeredSomething = 1
                FreeResponse = Session.Item("Cal_" & iAnswerCount.ToString)
            End If
        End If

        'put '?' in input fields if no value
        If FreeResponse = "" Or IsNothing(FreeResponse) Then
            FreeResponse = "?"
        End If

        strCell = "<TD style=""questionText"" class=""questionText"">"
        strCell = strCell & iDate.ToString & "<br>"
        strCell = strCell & "&nbsp;<input type=""text"" size=""2"" maxlength=""2"""

        If Not IsNothing(intMax) Then    'its a question that needs specific min/max validation
            strCell = strCell & " onBlur=""validateCalMinMax(this," & iMinimum.ToString & ", " & intMax.ToString & ")"""
        Else
            strCell = strCell & "  onBlur=""validate(this)"""
        End If
        strCell = strCell & " onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
        strCell = strCell & "value=""" & FreeResponse & """ id=""" & name & """ name=""" & name & """></td>" & vbCrLf

        strTable = strTable & strCell
    End Sub

    Private Sub addStaticField(ByRef strTable As String, ByVal iDate As Integer)
        Dim strCell As String
        strCell = "<TD valign=""top"">"
        If iDate > 0 Then
            strCell = strCell & iDate.ToString & "<br> &nbsp; <br></td>" & vbCrLf
        Else
            strCell = strCell & "&nbsp;<br> &nbsp; <br></td>" & vbCrLf
        End If


        strTable = strTable & strCell
    End Sub


    Sub storeTextAnswers(ByVal strTextQuestionID As String)
        'Dim cleanSQL() As String
        Dim strTextAnswer As String
        If Not IsNothing(Request.Form.Item(strTextQuestionID)) Then
            strTextAnswer = Trim(Request.Form.Item(strTextQuestionID))
            If (Not IsNothing(strTextAnswer)) Then    '1/3/01
                strTextAnswer = cleanSQL(strTextAnswer)
            End If
            Call addtoSession(strTextAnswer)
        End If
    End Sub

    Sub addtoSession(ByVal strTextAnswer As String)
        Dim strName As String
        Dim iCounter As Integer
        If Not IsNothing(Session.Item("calCounter")) Then
            iCounter = Session.Item("calCounter")
        Else
            iCounter = 0
        End If
        iCounter = iCounter + 1
        Session.Item("calCounter") = iCounter
        strName = "Cal_" & iCounter.ToString
        Session.Item(strName) = strTextAnswer
    End Sub

    Private Function prePopulateCal(ByVal iCounter As Integer) As String
        Dim strName As String
        Dim strAnswer As String
        strName = "Cal_" & iCounter.ToString
        If Not IsNothing(Session.Item(strName)) Then
            strAnswer = Session.Item(strName)
            answeredSomething = 1
        Else
            strAnswer = ""
        End If
        Return strAnswer
    End Function

    Private Function getHighestDrinkValue() As String
        'This function will get the highest drink value from the calendar grid and return the value, and the date 
        Dim strResult As String
        Dim dtDate As Date
        Dim i As Integer
        Dim iDrink As Integer = 0
        Dim strQuestion As String
        Dim strDate As String
        Dim arr(14, 1) As String
        Dim bAnswer As Boolean
        Dim ii As Integer

        If Not IsNothing(Session.Item("DateArray")) Then
            If IsArray(Session.Item("DateArray")) Then
                arr = Session.Item("DateArray")
                For i = 0 To UBound(arr)
                    ii = i + 1
                    If Not IsNothing(Session.Item("cal_" & ii.ToString)) And IsNumeric(Session.Item("cal_" & ii.ToString)) Then
                        If CInt(Session.Item("cal_" & ii.ToString)) >= iDrink Then
                            iDrink = CInt(Session.Item("cal_" & ii.ToString))
                            strQuestion = arr(i, 1)
                            strDate = arr(i, 0)
                        End If
                    End If
                Next
                If iDrink > 0 Then
                    bAnswer = True
                End If
            End If
        End If

        If Not bAnswer Then
            strResult = " not reported.  Would you consider returning to the previous page and entering values on the Calendar?</b>"
        Else
            strResult = iDrink.ToString & " drinks on " & getDay(strDate) & ", " & FormatDateTime(DateValue(strDate), DateFormat.ShortDate).ToString & ". On this particular day, over what time period were you drinking?</b>"
        End If

        Return "<span class=""highlight"">" & strResult & "</span>"
    End Function
    'Private Function getDDValues(ByRef arrStates() As String) As Array
    '    'fill array
    '    arrStates = New String() {"Choose a State", "Alabama", "Alaska", "American Samoa", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "District Of Columbia", "Federated States of Micronesia", "Florida", "Georgia", "Guam", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Marshall Islands", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Northern Mariana Islands", "Ohio", "Oklahoma", "Oregon", "Palau", "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virgin Islands", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", "Did not graduate in the United States"}
    'End Function

    '-- End of functions for calendar question type

    Private Function getDay(ByVal strDate As String) As String
        Dim dtDate As Date
        Dim strDay As String = ""
        Dim iDay As Integer
        If IsDate(strDate) Then
            dtDate = DateValue(strDate)
            'original line of code by heather
            'strDay = WeekdayName(dtDate.DayOfWeek)
            strDay = dtDate.DayOfWeek.ToString
        End If
        Return strDay
    End Function

    Private Function checkSkipForward2() As Integer
        Dim iTestFormID As Integer
        Dim iSkipToPage As Integer
        Dim objSkipPattern As New Navigation.SkipPattern
        iTestFormID = getTestFormID()
        iSkipToPage = objSkipPattern.checkSkipForward(pageNumber, iTestFormID)
        If iSkipToPage <> pageNumber Then
            Session.Add("sb" & CStr(iSkipToPage), pageNumber)
        End If
        Return iSkipToPage
    End Function

    Private Function checkSkipBack2() As Integer
        If Not IsNothing(Session.Item("sb" & CStr(pageNumber))) Then
            Return CInt(Session.Item("sb" & CStr(pageNumber)))
        Else
            Return pageNumber - 1
        End If

    End Function


    Private Sub RemoveCalSession()
        Dim i As Integer
        For i = 1 To 14
            If Not IsNothing(Session.Item("Cal_" & i)) Then
                Session.Remove("Cal_" & i)
            End If
        Next

        If Not IsNothing(Session.Item("DateArray")) Then
            Session.Remove("DateArray")
        End If
    End Sub

    Private Sub ClassifyV9Path()
        Dim ObjAssessment As New DataAccess.DAAssessment
        ObjAssessment.ClassifySurvey1(iUserAccountID)
        Dim dr As SqlClient.SqlDataReader
        Dim objectUsers As New DataAccess.DAUsers
        dr = objectUsers.SetUserClassification(iUserAccountID, 1)
    End Sub

    Private Sub moveNext()
        Dim objChangeSegment As Navigation.ChangeSegment = New Navigation.ChangeSegment
        strURL = objChangeSegment.getNextURL()
        If Trim(strURL) <> "" Then
            Response.Redirect(strURL)
        Else
            Response.Redirect("/error.aspx")
        End If
    End Sub

    Private Function editForFemale(ByVal strQuestion As String, ByVal bFemale As Boolean) As String
        Dim newQuestion As String
        If bFemale Then
            newQuestion = Replace(strQuestion, "5", "4")
            newQuestion = Replace(newQuestion, "five", "four")
        Else
            newQuestion = strQuestion
        End If
        Return newQuestion
    End Function

    Private Sub setSurveyGender(ByVal strChoice As String)
        If strChoice = "25853" Then
            Session.Item("SurveyFemale") = True
        Else
            Session.Item("SurveyFemale") = False
        End If

    End Sub

    Private Function getUserClassification(ByVal iClassificationID As Integer) As String
        Dim objClassification As DataAccess.DAUsers = New DataAccess.DAUsers
        Return objClassification.checkClassification(iUserAccountID, iClassificationID)
    End Function
End Class
