Imports System.Data
Imports AlcoholEdu
Imports AlcoholEdu.DataAccess
Imports System.Web.UI.HtmlControls

Public Class evaluation
	Inherits System.Web.UI.Page

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

    Dim lc_page_type As String = "/survey/evaluation.aspx"
	Dim iUserAccountID As Integer
	Dim iAssessmentID As Integer
	Dim iSegmentID As Integer
    Dim bSOC As Boolean
    Dim bSETwo As Boolean
    Dim bSEOne As Boolean
    Dim iPageNumber As Integer
    Dim bAnswered As Boolean
    Dim intlastPage As Integer
    Dim intNoFields As Integer
    Dim intPageStart As Integer
    Dim strAllPopUpDivs As String

    'Protected WithEvents PopUps As System.Web.UI.HtmlControls.HtmlGenericControl

	Const qt17 As Integer = 1
	Const qt18 As Integer = 2
	Const qt19 As Integer = 3
	Const qt20 As Integer = 4
	Const qt21 As Integer = 5
	Const qt22 As Integer = 6
	Const qt23 As Integer = 7
	Const qt40 As Integer = 14
	Const qt41 As Integer = 15
	Const qt60 As Integer = 16
	Const qt45 As Integer = 8
	Const qt46 As Integer = 9
	Const qt48 As Integer = 11
	Const qt54 As Integer = 10
	Const qt86 As Integer = 12
	Const qt144 As Integer = 13
	Const qf24 As Integer = 1
	Const qf25 As Integer = 2
	Const qf26 As Integer = 3
	Const qf27 As Integer = 4
	Const qf49 As Integer = 5
	Const qf50 As Integer = 6
	Const qf51 As Integer = 7
    Const qf143 As Integer = 8
    Const qf999 As Integer = 16
    Dim strMsg As String = ""

    Protected WithEvents subTitle As Label
    Protected WithEvents surveyImage As Image

	Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		Dim i As Integer
		Dim answeredSomething As Byte
		Dim lastPage As String
		Dim Format_Renamed As Short
		Dim pageNumber As String
		Dim intAnswerFlag As Integer
		Dim iTestTypeID As Integer
		Dim strReqQuestions As String
		Dim iClientID As Integer
		Dim strSegmentName As String

		Dim strSkipQuestion As String
        Dim iSectionTypeID As Integer
        Dim iTestFormID As Integer
		Dim SegmentType As Integer
		Dim strHTMLQuestions As String
		Dim strHTMLButton As String

        Const formatFORMFLAG As Short = 1
        Dim contentPlaceHolder As New ContentPlaceHolder
        Dim myControl As New HtmlGenericControl
        Dim bErrors As Boolean = False
		iSegmentID = Session.Item("iSegmentID")
		iUserAccountID = Session.Item("iUserAccountID")

        setTestInformation(iUserAccountID, iSegmentID)
        Trace.Write("iSectionTypeID: " & Session.Item("iTestID"))
        intAnswerFlag = 0
        'iClientID = Session.Item("iClientID")
        strSegmentName = Session.Item("strSegmentName")
        iSectionTypeID = Session.Item("iTestID")
        iTestFormID = Session.Item("iTestFormID")
        'IF the user is in the wrong SegmentType, send them packing
        Trace.Write("iSectionTypeID: " & iSectionTypeID)

        Select Case iSectionTypeID
            Case 22
                Trace.Write("Passed")
            Case Else
                'Response.End()
                Response.Redirect("/error.aspx?msg=back")
        End Select

        '		pageNumber = 0
        Format_Renamed = formatFORMFLAG
        strReqQuestions = ""

        iAssessmentID = Session.Item("iAssessmentID")

        'If Request.Item("assessmentID") <> "" And CDbl(Request.Item("assessmentID")) <> 0 Then
        'iAssessmentID = Request.Item("assessmentID")
        'End If

        If Not IsNothing(Request.Form.GetValues("pageNumber")) Then pageNumber = Request.Item("pageNumber")
        If IsNumeric(pageNumber) Then
            iPageNumber = pageNumber
        Else
            iPageNumber = 0
        End If

        ' iTestTypeID = GetTestTypeID(iSegmentID)

        iPageNumber = iPageNumber + 1

        'Dim objButton2 As HtmlInputButton
        'objButton2 = FindControl("button2")
        'objButton2.Visible = False


        If Not Page.IsPostBack And iPageNumber < 2 Then
            'Before we do anything else, let's set the pageheader propery for this page
            'Dim myHeaderControl As headerNoNav
            'myHeaderControl = Page.FindControl("pageHeader")
            'myHeaderControl.lc_page_type = lc_page_type
            'Dim myFooterControl As PageFooter
            'myFooterControl = Page.FindControl("pageFooter")
            'myFooterControl.lc_page_type = lc_page_type

            'pageNumber = 1

            If bSOC Then
                'Get the IntroText for above the questions
                strMsg = getIntroText(iClientID)
                Dim objButton As HtmlInputButton
                objButton = FindControl("button1")
                objButton.Visible = False
                'objButton2.Visible = True

            End If

            If iTestFormID = 2 Then
                strAllPopUpDivs = ""
                'subTitle.text = "<h2>Institutionalization Diagnostic</h2>"
                surveyImage.ImageUrl = "/images/Survey2.png"
            End If

            strHTMLQuestions = getEvaluationQuestions(bSOC, iClientID, iTestFormID)
            'output questions for display    
            'Dim objInterviewQuestions As HtmlGenericControl
            'objInterviewQuestions = FindControl("evaluationQuestions")
            'objInterviewQuestions.InnerHtml = strHTMLQuestions

            contentPlaceHolder = Master.FindControl("mainContent")

            strHTMLQuestions = strAllPopUpDivs & strHTMLQuestions

            myControl = contentPlaceHolder.FindControl("evaluationQuestions")
            myControl.InnerHtml = strHTMLQuestions

            'PopUps = FindControl("PopUps")
            'PopUps.InnerText = strAllPopUpDivs
            'Dim alldivtags As Literal
            'alldivtags = FindControl("alldivtags")
            'If strAllPopUpDivs <> "" Then
            '    alldivtags.text = strAllPopUpDivs

            'End If


        Else    'on postback
            Call saveSurvey()
            If isValidSurvey(iSegmentID) Then


                If bSOC And iPageNumber <> 3 Then           'Check for skip question
                    Dim bClassify As Boolean
                    bClassify = classifySOC()

                    strSkipQuestion = Request.Form("a4574")
                    If strSkipQuestion = "16875" Or strSkipQuestion = "16876" Then
                        iPageNumber = 3
                    Else
                        'Display the third question
                        strMsg = ""
                        'objButton2.Visible = True
                        strHTMLQuestions = getEvaluationQuestions(bSOC, iClientID, iTestFormID)
                        'output questions for display    
                        'Dim objInterviewQuestions As HtmlGenericControl
                        'objInterviewQuestions = FindControl("evaluationQuestions")
                        'objInterviewQuestions.InnerHtml = strHTMLQuestions
                        'Dim contentPlaceHolder As New ContentPlaceHolder
                        contentPlaceHolder = Master.FindControl("mainContent")
                        'Dim myControl As HtmlGenericControl
                        myControl = contentPlaceHolder.FindControl("evaluationQuestions")
                        myControl.InnerHtml = strHTMLQuestions

                    End If
                End If
            Else
                'Show the error and stop processing
                'Could just move the getEvaluationQuestionStuff here....
                Dim strErrorMsg As String
                strErrorMsg = "<b>Please correct the following error(s) that are present in this page:</b>"
                strMsg = strErrorMsg & strMsg
                bErrors = True
            End If


            If (bSOC = False Or iPageNumber = 3) And (bErrors = False) Then
                'set dtEnd in assessment
                Call setSurveyDate()
                Session.Item("NavFunction") = "next"

                'Move to next Segment
                Dim strURL
                'Dim objChangeSegment = New Navigation.ChangeSegment
                'strURL = objChangeSegment.getNextURL()
                strURL = getNextPage()
                If Trim(strURL) <> "" Then
                    Response.Redirect(strURL)
                Else
                    Response.Redirect("/error.aspx")
                End If
            End If

        End If       'End check for postback


        'Dim objIntroText As HtmlGenericControl
        'objIntroText = FindControl("errorMessage")
        'objIntroText.InnerHtml = strMsg
        'Dim contentPlaceHolder As New ContentPlaceHolder
        contentPlaceHolder = Master.FindControl("mainContent")
        'Dim myControl As HtmlGenericControl
        myControl = contentPlaceHolder.FindControl("errorMessage")
        myControl.InnerHtml = strMsg
        myControl.DataBind()

        'Page.DataBind()


    End Sub

    Private Function getNextPage() As String
        'TODO: not sure if we are going to feedback before the second survey or showing a choice...
        'Return "feedback.aspx"
        If Session.Item("iSegmentID") = 1 Then
            Session.Item("iSegmentID") = 2
            Return "evaluation.aspx"
        Else
            Return "feedback.aspx"
        End If
    End Function
    Function getEvaluationQuestions(ByVal bSOC As Boolean, ByVal iClientID As Integer, ByVal iTestFOrmID As Integer) As String
        Dim DRInterview As DataSet
        Dim strHtml As String
        Dim intQuestionID As Integer
        Dim intQuestionFormatID As Integer
        Dim strQuestions As String
        Dim intQuestionTypeID As Integer
        Dim intQuestionNumber As Integer
        Dim intHoldDataViewQuestionID As Integer
        Dim intAccessKey As Integer
        Dim strAccessKey As String
        Dim QuestionTable As DataTable
        Dim QuestionColumn As DataColumn
        Dim QuestionRow As DataRow
        Dim i As Integer
        Dim j As Integer = 0
        Dim x As Integer = 0
        Dim intDataViewIncrement As Integer
        Dim intDataViewQuestionID As Integer
        Dim FreeResponse As String
        Dim responseSelected As Boolean
        Dim strTextQuestion As String
        Dim answeredSomething As Boolean
        Dim intMax As Integer
        Dim intMin As Integer
        Dim strSubQuestions As String
        Dim strReqQuestions As String
        Dim bSpecialFormat As Boolean
        Dim strStyle As String
        Dim intChoicesCount As Integer
        Dim intIdcounter As Integer
        Dim intNoChoices As Integer
        Dim strQuestionText As String
        Dim strQuestionChoice As String
        'Dim HTML2 as String 'This is the question for the second page

        'TODO: Add code to skip the last question on SOC based on answers to other questions
        'HOW Do I know we are in SOC vs OO, they both have the same SectionType/TestType. Go by number of questions? OR associate with questionTypeID
        'get the list values

        'Dim objInterview As New Content.PopulateLists
        'DRInterview = objInterview.getQuestions(iSegmentID, iAssessmentID, iPageNumber)
        Dim objQuestions As New AlcoholEdu.DataAccess.DAContent
        Dim ds As DataSet
        If bSOC Then
            iAssessmentID = Session.Item("iAssessmentID")
            ds = objQuestions.GetQuestions(iTestFOrmID, iAssessmentID, 1)
        Else
            ds = objQuestions.GetQuestionsForXML(iTestFOrmID, 1)
        End If



        ''Start Table to display
        'If Not bSOC Then
        'strHtml = "<table width=""100%"" cellpadding=0 cellspacing=0 border=0>"
        'strHtml = strHtml & "<tr><td width=70>&nbsp;</td>"
        'strHtml = strHtml & "<td width=""15%"" align=right><img height=15 width=99 src=""/images/previous_page_grey.gif"" border=0 alt=""Previous Page""></td>"
        'strHtml = strHtml & "<td width=""15%""><img height=15 width=77 src=""/images/next_page_grey.gif"" border=0 alt=""Next Page""></td></tr>"
        'strHtml = strHtml & "</table><br>"
        'End If
        Dim iRunningCount As Integer = 0
        strHtml = strHtml & "<table width=""750"" cellspacing=0 cellpadding=0 border=0><tr><td>"
        'Dim QuestionTable As DataTable
        'all the code for setting up and formatting the questions will go here
        If ds.Tables.Count > 0 Then          'dataset contains a table...proceed
            QuestionTable = ds.Tables(0)

            'For Each QuestionTable In DRInterview.Tables
            strHtml = strHtml & "<table width=""100%"" cellspacing=0 cellpadding=1 border=0>"

            'all the code for setting up and formatting the questions will go here
            For Each QuestionRow In QuestionTable.Rows
                'loop through all survey questions
                'intIOrder = QuestionRow.Item("iOrder")
                'If intIOrder > 5000 Then    'custom questions all start with 5000/Qset questions all start with 6000
                '    blnCustomQuestion = True
                'End If
                intQuestionID = QuestionRow.Item("iQuestionID")
                intQuestionFormatID = QuestionRow.Item("iQuestionFormatID")
                intQuestionTypeID = QuestionRow.Item("iQuestionTypeID")
                intQuestionNumber = QuestionRow.Item("iQuestionNumber")
                intlastPage = QuestionRow.Item("iLastPage")
                intPageStart = QuestionRow.Item("iPageStart")
                bAnswered = False
                responseSelected = False
                'HRHHERE
                iRunningCount = iRunningCount + 1

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
                                'intCustomQuestionNumber = intCustomQuestionNumber + 1
                                strHtml = strHtml & "<tr>" & vbCrLf
                                strHtml = strHtml & "<td class=""questionText"" style=""questionText""><b>"
                                'If blnCustomQuestion And intIOrder < 6000 Then       'its a custom question (and not a Qset question) so we need to use dynamic numbering
                                '    strHtml = strHtml & intCustomQuestionNumber & ".&nbsp;&nbsp;</b>"
                                'Else
                                strHtml = strHtml & QuestionRow.Item("iQuestionNumber") & ".&nbsp;&nbsp;</b>"
                                'End If
                                If intQuestionTypeID = 10 Then
                                    If InStr(QuestionRow.Item("txQuestion"), "<<DrinkQTY>>") Then
                                        strHtml = strHtml & Replace(QuestionRow.Item("txQuestion"), "<<DrinkQTY>>", 6)
                                    Else
                                        strHtml = strHtml & QuestionRow.Item("txQuestion")
                                    End If
                                Else
                                    'If intQuestionID = 6246 Or intQuestionID = 6256 Then
                                    '    'These have different text if female.
                                    '    strHtml = strHtml & editForFemale(QuestionRow.Item("txQuestion"), bFemale)
                                    'Else
                                    If QuestionRow.Item("tiAnswerable") = 0 Then
                                        strHtml = strHtml & "<div class=""highlight"">" & QuestionRow.Item("txQuestion") & "</div>"
                                    Else
                                        strHtml = strHtml & QuestionRow.Item("txQuestion")
                                    End If

                                    'End If

                                End If
                                strHtml = strHtml & "</td>" & vbCrLf
                                strHtml = strHtml & "</tr>" & vbCrLf
                            Else
                                strHtml = strHtml & "<tr>" & vbCrLf
                                strHtml = strHtml & "<td class=""questionText"" style=""questionText"">"

                                If QuestionRow.Item("tiAnswerable") = 0 Then
                                    strHtml = strHtml & "<div class=""highlight"">" & QuestionRow.Item("txQuestion") & "</div>"
                                Else
                                    strHtml = strHtml & QuestionRow.Item("txQuestion")
                                End If

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
                                    'bCalQuestion = True
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
                                        'strHtml = strHtml & createCalendarGrid(QuestionChoiceDataView, dvAnswers, intQuestionID, intDataViewIncrement, intHoldDataViewQuestionID, strReqQuestions, strTextQuestion, strHiddenFieldValues)
                                        'create the hidden form fields
                                        strHtml = strHtml & "<input type=""hidden"" name=""a" & intQuestionID & "_MC"" value=""" & strHiddenFieldValues & """>" & vbCrLf
                                        'put a hidden form field so we can tell when we're saving answers that we need to store these in the session too.
                                        strHtml = strHtml & "<input type=""hidden"" name=""calQuestion"" value=""" & intQuestionID & """>" & vbCrLf
                                        'Move on
                                    End If

                                Case 6, 1, 3          'Vertical, MCSR, MCMR, MFR
                                    Dim strIDForPopUp As String = ""
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
                                        'responseSelected = False 
                                        If bSOC Then
                                            FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                            responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                        Else
                                            responseSelected = False
                                        End If
                                        bAnswered = responseSelected
                                        intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                        If intDataViewQuestionID = intQuestionID Then
                                            intDataViewIncrement = intDataViewIncrement + 1
                                            intHoldDataViewQuestionID = intDataViewQuestionID
                                            'Now check for this answer in the dvAnswers DataView
                                            If intQuestionTypeID = 6 Then             'Only allow the Open Text questsions to use this one.
   
                                                'HRH 2/7/07 Adding this to not show pop-up if the free response is the only question on the page
                                                'If QuestionChoiceDataView.Count = 1 Then
                                                '    'There is only one question on the page
                                                '    answeredSomething = 1
                                                'End If
                                                'Code to track how often this is called.
                                                'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 1")
                                            End If
                                            If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = True
                                            Else
                                                'FreeResponse = ""
                                                'responseSelected = False
                                            End If
                                            strHtml = strHtml & "<tr>"

                                            If intQuestionFormatID = 4 Then             ' LARGE Question (or Text Areas)
                                                strHtml = strHtml & "<td class=answerNumber>" & Chr(13)
                                                strHtml = strHtml & "<textarea "
                                                strHtml = strHtml & "onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"""
                                                strHtml = strHtml & " name=""a" & intQuestionID & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")
                                                strHtml = strHtml & " rows=5 cols=40 onKeyUp=""return maxCharaters(this)"">" & FreeResponse & "</textarea>" & Chr(13)
                                                'strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """>" & Chr(13)
                                                strTextQuestion = strTextQuestion & "," & intQuestionID
                                            Else
                                                strHtml = strHtml & "<td valign=top width=25 class=""answerNumber1"" style=""answerNumber1"">" & Chr(13)
                                                'HRHHERE2
                                                If QuestionChoiceDataView.Item(i).Row("iSubQuestionID") > 0 Then
                                                    If Trim(QuestionChoiceDataView.Item(i).Row("chSelectionNumber")) = "o" Then
                                                        'Don't do the pop-up on the "other" question even though it has a sub-question.
                                                        strIDForPopUp = ""
                                                    Else
                                                        strIDForPopUp = QuestionChoiceDataView.Item(i).Row("iSubQuestionID")
                                                    End If

                                                Else
                                                    strIDForPopUp = ""
                                                End If

                                                strHtml = strHtml & "<input onClick=""setAnsweredSomething(" & strIDForPopUp & ");""  type="

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
                                                bAnswered = responseSelected 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
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
                                                Trace.Write("QuestionFormatInfo", CStr(intQuestionID) & " " & QuestionRow.Item(17))
                                                strHtml = strHtml & QuestionRow.Item("chSelectionNumber")
                                            End If
                                            If Not (IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice"))) Then
                                                strHtml = strHtml & "</td><td valign=top class=""answerText"" style=""answerText"">"
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")
                                            End If
                                            If QuestionChoiceDataView.Item(i).Row("iSubQuestionID") <> 0 Then
                                                strSubQuestions = getSubQuestion(QuestionChoiceDataView.Item(i).Row("iSubQuestionID"), QuestionTable.Copy)
                                                strHtml = strHtml & strSubQuestions
                                                'will need to increment for the sub-question
                                                strSubQuestions = ""
                                                intDataViewIncrement = intDataViewIncrement + 1
                                            End If
                                            strHtml = strHtml & "</td></tr>" & Chr(13)
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
                                            responseSelected = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 3")
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                responseSelected = False
                                            End If
                                            If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = True
                                            Else
                                                FreeResponse = ""
                                                'responseSelected = False
                                            End If
                                            strHtml = strHtml & "<td valign=bottom width=""10%"" class=""answerNumber"" style=""answerNumber"">" & Chr(13)
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
                                                bAnswered = responseSelected 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
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
                                                    'Trace.Write("QuestionFormatInfo", CStr(iQuestionID) & " " & QuestionRow.Item(17))
                                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>" & Chr(13)
                                                End If
                                            Else
                                                strHtml = strHtml & "</td><td class=""answerText"" style=""answerText"">" & Chr(13)
                                            End If
                                            If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")) Then

                                                If Not ((intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7)) Then
                                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "" & Chr(13)
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
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                responseSelected = False
                                            End If

                                            'responseSelected = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                            strHtml = strHtml & "<td valign=bottom width=""10%"" class=""answerNumber"" style=""answerNumber"">" & Chr(13)
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
                                                bAnswered = responseSelected 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
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
                                                    ' Trace.Write("QuestionFormatInfo", CStr(iQuestionID) & " " & QuestionRow.Item(17))
                                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>" & Chr(13)
                                                End If
                                            Else
                                                strHtml = strHtml & "</td><td class=""answerText"" style=""answerText"">" & Chr(13)
                                            End If
                                            If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice")) Then

                                                If Not ((intQuestionFormatID <> 3 Or intQuestionFormatID <> 4) And (intQuestionTypeID <> 7)) Then
                                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "" & Chr(13)
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
                                            responseSelected = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 5")
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                responseSelected = False
                                            End If

                                            If QuestionChoiceDataView.Item(i).Row("vchFreeResponse") <> "" Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = True
                                            Else
                                                FreeResponse = ""
                                                'responseSelected = False
                                            End If
                                            'FreeResponse = ""
                                            Dim strExtrNote As String = ""
                                            strHtml = strHtml & "" & Chr(13)
                                            If intDataViewQuestionID = 32 Then
                                                strHtml = strHtml & "<b>$ &nbsp;</b>"
                                            ElseIf intDataViewQuestionID = 31 Then
                                                'If FreeResponse = "" Then FreeResponse = "0.0"
                                                strExtrNote = " <B>FTE</B> (Please enter a numeric value, for example 3.5)"
                                            End If
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

                                            '=============================================================================================================================================
                                            strHtml = strHtml & " name=""a" & intQuestionID
                                            strHtml = strHtml & "text" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ value=""" & FreeResponse & """ >"

                                            If QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") <> "" Then
                                                strHtml = strHtml & "&nbsp;" & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & Chr(13)
                                            End If
                                            If FreeResponse <> "" And Not IsDBNull(FreeResponse) Then
                                                answeredSomething = 1
                                            End If
                                            strHtml = strHtml & strExtrNote
                                            strHtml = strHtml & "<input type=hidden name=""a" & intQuestionID & "_MC" & """ value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            strTextQuestion = strTextQuestion & "," & intQuestionID
                                            'strHtml = strHtml & "&nbsp; &nbsp;" & Chr(13)
                                        Else
                                            'finish row before exiting...
                                            'strHtml = strHtml & "<tr><td>&nbsp;</td></tr>" & Chr(13)
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
                                        strHtml = strHtml & "<table width=750 cellpadding=0 cellspacing=0 class=""gridTable"" border=0>" & Chr(13)
                                    End If
                                    strStyle = "surveyQuestionText"
                                    If CBool(j Mod 2) Then
                                        strStyle = strStyle & "1"
                                    End If
                                    'Create a new dataview instance on the table to iterate through question choices
                                    Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)

                                    strHtml = strHtml & "<tr><td class=""" & strStyle & """ style=""" & strStyle & """"
                                    strHtml = strHtml & " width=400>" '& QuestionChoiceDataView.Item(i).Row("iQColumnWidth") & ">"
                                    'If intQuestionID = 6246 Or intQuestionID = 6256 Then
                                    '    'These have different text if female.
                                    '    strHtml = strHtml & editForFemale(QuestionChoiceDataView.Item(i).Row("txQuestion"), bFemale) & Chr(13)
                                    'Else
                                    If CBool(QuestionRow.Item("tiDisplayQuestionNumber")) Then
                                        strHtml = strHtml & QuestionRow.Item("iQuestionNumber") & ".&nbsp;&nbsp;"
                                    End If
                                    ' strHtml = strHtml & QuestionRow.Item("iQuestionNumber") & ".&nbsp;&nbsp;</b>"
                                    strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                    'End If
                                    'strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("txQuestion") & Chr(13)
                                    strHtml = strHtml & "</td>"
                                    'HRHHERE
                                    intDataViewIncrement = iRunningCount
                                    For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                        intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")

                                        If intDataViewQuestionID = intQuestionID Then
                                            intDataViewIncrement = intDataViewIncrement + 1
                                            intHoldDataViewQuestionID = intDataViewQuestionID
                                            intChoicesCount = intChoicesCount + 1
                                            strHtml = strHtml & "<td  class=""" & strStyle & """ width=100" '& QuestionChoiceDataView.Item(i).Row("iColumnWidth") & "
                                            strHtml = strHtml & " style=""text-align:center;"">" & Chr(13)

                                            If intQuestionFormatID = 17 Then 'for dropdown
                                                strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                                'js 4/14/09: for the new grid-style-checkbox...
                                                'ElseIf intQuestionFormatID = 18 Then 'for checkboxes
                                                'strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=checkbox"
                                            Else 'for radio buttons or anything not above
                                                strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                            End If
                                            If Not IsDBNull(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) Then
                                                'bAnswered = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                                If bSOC Then
                                                    FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                    bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                                Else
                                                    bAnswered = False
                                                End If
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
                                                    'If vchDropDown = CStr(a) Then
                                                    '    strSelected = "selected"
                                                    'Else
                                                    strSelected = ""
                                                    'End If
                                                    strHtml = strHtml & "<option " & strSelected & ">" & a & "</option>"
                                                Next

                                                strHtml = strHtml & "</select>"
                                                strHtml = strHtml & "</td>" & Chr(13)
                                                strHtml = strHtml & "<input type=hidden name=dropdown" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Else 'any format not dropdown...radio or checkbox
                                                strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></td>" & Chr(13)
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
                                    strHtml = strHtml & "<table width=750 cellpadding=0 cellspacing=0 border=0>" & Chr(13)
                                    strStyle = "surveyQuestionText"
                                    If CBool(j Mod 2) Then
                                        strStyle = strStyle & "1"
                                    End If

                                    intDataViewIncrement = iRunningCount
                                    'Create a new dataview instance on the table to iterate through question choices
                                    Dim QuestionChoiceDataView As DataView = New DataView(QuestionTable)

                                    strHtml = strHtml & "<tr><td class=""" & strStyle & """ style=""" & strStyle & """"
                                    strHtml = strHtml & " width=400>" ' & QuestionChoiceDataView.Item(i).Row("iQColumnWidth") & ">"
                                    If CBool(QuestionRow.Item("tiDisplayQuestionNumber")) Then
                                        strHtml = strHtml & QuestionRow.Item("iQuestionNumber") & ".&nbsp;&nbsp;"
                                    End If
                                    strHtml = strHtml & QuestionRow.Item("txQuestion") & Chr(13)
                                    strHtml = strHtml & "&nbsp;&nbsp;</td>"
                                    For i = intDataViewIncrement To QuestionChoiceDataView.Count - 1
                                        intDataViewQuestionID = QuestionChoiceDataView.Item(i).Row("iQuestionID")
                                        If intDataViewQuestionID = intQuestionID Then


                                            strHtml = strHtml & "<td class=""" & strStyle & """ width=100 " '& QuestionChoiceDataView.Item(i).Row("iColumnWidth") &
                                            strHtml = strHtml & "style=""text-align:center;"">" & Chr(13)
                                            If Not (IsDBNull(QuestionChoiceDataView.Item(i).Row("vchQuestionChoice"))) Then
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>"
                                            Else
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & " <br> """
                                            End If

                                            If intQuestionFormatID = 17 Then 'for dropdown
                                                strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                                'ElseIf intQuestionFormatID = 18 Then 'for checkbox
                                                'strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=checkbox"
                                            Else 'for radio + any format not above
                                                strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                            End If

                                            'bAnswered = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                bAnswered = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                bAnswered = False
                                            End If

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
                                                    'If vchDropDown = CStr(a) Then
                                                    '    strSelected = "selected"
                                                    'Else
                                                    strSelected = ""
                                                    'End If
                                                    strHtml = strHtml & "<option " & strSelected & ">" & a & "</option>"
                                                Next

                                                strHtml = strHtml & "</select>"
                                                strHtml = strHtml & "</td>" & Chr(13)
                                                strHtml = strHtml & "<input type=hidden name=dropdown" & intQuestionID & " value=" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Else
                                                strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></td>" & Chr(13)
                                            End If
                                            intDataViewIncrement = intDataViewIncrement + 1
                                            intHoldDataViewQuestionID = intDataViewQuestionID
                                            intIdcounter = intIdcounter + 1
                                            intNoChoices = intNoChoices + 1
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
                                            'responseSelected = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, dvAnswers)
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                responseSelected = False
                                            End If
                                            strHtml = strHtml & "" & Chr(13)
                                            strHtml = strHtml & "<select onChange=""setAnsweredSomething();"" "
                                            'bAnswered = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            bAnswered = responseSelected
                                            If bAnswered Then
                                                answeredSomething = 1
                                            End If
                                            strHtml = strHtml & " name=""a" & intQuestionID & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & ">" & Chr(13)
                                            Dim a As Integer

                                            '=============NEW CODE FOR HANDLING MULTIPLE DROPDOWNS WITH A SINGLE SURVEY=====================================
                                            'JS:all code between '===' is new code added 6/14 to handle multiple drop downs within one survey 
                                            'will proceed with the concept of using the hardcoded questionIDs of the drop down type questions...not the best way to
                                            '  go but have no time to do it more elegantly
                                            Dim arrDropDown() As String
                                            Dim strTypeArray As String
                                            Dim strExtraNote As String = ""
                                            Select intQuestionID
                                                Case 32 'Budget DropDown
                                                    strTypeArray = "budget"
                                                    arrDropDown = makeArray(strTypeArray)
                                                Case Else
                                                    strTypeArray = "error"
                                                    arrDropDown = makeArray(strTypeArray)
                                            End Select

                                            Dim strSelected As String
                                            a = 0
                                            For a = a To UBound(arrDropDown)
                                                'If vchDropDown = CStr(arrDropDown(a)) Then
                                                '    strSelected = "selected"
                                                'Else
                                                strSelected = ""
                                                'End If
                                                strHtml = strHtml & "<option " & strSelected & " value=""" & a.ToString & """>" & arrDropDown(a) & "</option>"
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
                                            strHtml = strHtml & strExtraNote & "</td>" & Chr(13)
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
                                            strHtml = strHtml & ">" & Chr(13)
                                            If Trim(QuestionChoiceDataView.Item(i).Row("chSelectionNumber")) <> "" Then
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("chSelectionNumber") & "<br>"
                                            Else
                                                strHtml = strHtml & QuestionChoiceDataView.Item(i).Row("vchQuestionChoice") & "<br>"
                                            End If
                                            strHtml = strHtml & "<input onClick=""setAnsweredSomething();"" onKeyPress=""setAnsweredSomething();"" type=radio"
                                            If bSOC Then
                                                FreeResponse = QuestionChoiceDataView.Item(i).Row("vchFreeResponse")
                                                responseSelected = checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), FreeResponse, QuestionChoiceDataView.Item(i).Row("iResponse"))
                                            Else
                                                responseSelected = False
                                            End If

                                            'bAnswered = False 'checkForAnswer(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID"), vchFreeResponse, dvAnswers)
                                            'Code to track how often this is called.
                                            'Trace.Write("CheckForAnswer ", CStr(QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID")) & " Location 8")
                                            bAnswered = responseSelected
                                            If bAnswered Then
                                                strHtml = strHtml & " checked"
                                                answeredSomething = 1
                                            End If
                                            'If QuestionChoiceDataView.Item(i).Row("iResponse") = QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") Then
                                            'strHtml = strHtml & " checked"
                                            'answeredSomething = 1
                                            'End If

                                            strHtml = strHtml & " name=""a" & intQuestionID & """ value=""" & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & """ id=id" & intIdcounter & QuestionChoiceDataView.Item(i).Row("iQuestionChoiceID") & "></td>" & Chr(13)
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
                                'bCalQuestion = True
                            End If
                        End If       'end bAnswerable = true
                    End If       'end bSubQuestion = 0
                End If    'end intHoldDataViewQuestionID <> intQuestionID
            Next    'QuestionRow In QuestionTable.Rows

            'If (pageNumber > 0 Or blnCustomQuestion) And (Not bCalQuestion) Then    'the blnCustomQuestion is to ensure that the page doesnt blow out when printing just custom questions
            '    strHtml = strHtml & "</table></table>"
            'Else
            strHtml = strHtml & "</td></tr></table>"
            'End If

            ' Next
        End If

        'print back and next links if not intro page
        'this will be pre-pended to the question string below
        'If intPageStart <> 1 Then
        '    strLinks = "<table width=""100%"" cellpadding=0 cellspacing=0 border=0>"
        '    strLinks = strLinks & "<tr><td width=70>&nbsp;</td>"
        '    If iPageNumber > 1 Then
        '        strLinks = strLinks & "<td width=""15%"" align=right><a href=""javascript:goToPreviousPage();"" AccessKey=""A""><img height=15 width=99 src=/images/previous_page.gif border=0 alt=""Previous Page""></a></td>"
        '    Else    'page = 1
        '        strLinks = strLinks & "<td width=""15%"" align=right><img height=15 width=99 src=""/images/previous_page_grey.gif"" border=0 alt=""Previous Page""></td>"
        '    End If
        '    If intlastPage <> 0 Then
        '        strLinks = strLinks & "<td width=""15%""><a href=""javascript:goToNextPage();"" AccessKey=B><img height=15 width=77 src=""/images/next_page.gif"" border=0 alt=""Next Page""></a></td>"
        '    Else
        '        strLinks = strLinks & "<td width=""15%""><img height=15 width=77 src=""/images/next_page_grey.gif"" border=0 alt=""Next Page""></td>"
        '    End If
        '    strLinks = strLinks & "</table><br>"
        'End If

        'strHtml = strHtml

        'Hidden field containing all Question IDs on this page.
        'strHtml = strHtml & "<input type=hidden value=""" & strReqQuestions & """ name=""reqQuestions"">" & _
        '  "<input type=hidden value=""" & strQuestions & """ name=""ansQuestions"">" & _
        '  "<input type=hidden value=""" & strTextQuestion & """ name=""ansTextQuestions"">" & _
        '  "<input type=hidden value=""" & answeredSomething & """ name=""answeredFlag"">" & _
        '  "<input type=hidden value=""" & navFunction & """ name=""navFunction"">" & _
        '  "<input type=hidden value=""" & iPageNumber & """ name=""pageNumber"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern1 & """ name=""blnSkipPattern1"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern2 & """ name=""blnSkipPattern2"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern3 & """ name=""blnSkipPattern3"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern4 & """ name=""blnSkipPattern4"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern5 & """ name=""blnSkipPattern5"">" & _
        '  "<input type=hidden value=""" & blnSkipPattern6 & """ name=""blnSkipPattern6"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern7 & """ name=""blnSkipPattern7"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern8 & """ name=""blnSkipPattern8"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern9 & """ name=""blnSkipPattern9"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern10 & """ name=""blnSkipPattern10"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern11 & """ name=""blnSkipPattern11"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern12 & """ name=""blnSkipPattern12"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern13 & """ name=""blnSkipPattern13"">" & _
        '    "<input type=hidden value=""" & blnSkipPattern14 & """ name=""blnSkipPattern14"">"

        If intlastPage = 0 Then
            strHtml = strHtml & "<input type=hidden value=""1"" name=""EndSurvey"">"
        End If

        'TODO: Dispose of that datatable


        Return strHtml


    End Function
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
        Dim intPos As Object
        Dim strResponses As String = ""
        Dim theAnswerArray() As String
        arrQuestionArray = Split(Request.Form.Item("ansQuestions"), ",")
        arrTextAnswerArray = Split(Request.Form.Item("ansTextQuestions"), ",")

        Dim x As Integer
        Dim iLen, iMid As Integer

        For x = 0 To Request.Form.Count - 1
            'Response.Write(Request.Form.Key(x) & " = ")
            'Response.Write(Request.Form.Item(x) & "<br>")
            'strResponses = strResponses & "," & Request.Form.Keys.Item(x)
            If Not IsNothing(Request.Form.Item(x)) Then
                'Check to see if it's a multipleChoice or text field
                If InStr(Request.Form.Keys.Item(x), "text") Then
                    strQuestionID = Request.Form.Keys.Item(x) 'a23text159 
                    iLen = Len(strQuestionID)
                    'iMid = InStr(strQuestionID, "text")
                    intQuestionChoice = CInt(Right(strQuestionID, iLen - InStrRev(strQuestionID, "t")))
                    strQuestionID = Left(strQuestionID, InStr(strQuestionID, "t") - 1)
                    strQuestionID = Replace(strQuestionID, "a", "")
                    intQuestionID = CInt(Trim(strQuestionID))
                    saveTextAnswers(intQuestionID, intQuestionChoice, Request.Form.Keys.Item(x))
                ElseIf InStr(Request.Form.Keys.Item(x), "_MC") Then
                    'If InStr(Request.Form.Keys.Item(x), "_MC") Then
                    'Loop through all selected choices and save them
                    strQuestionID = Request.Form.Keys.Item(x)
                    strQuestionID = Replace(strQuestionID, "_MC", "")
                    strQuestionID = Replace(strQuestionID, "a", "")
                    intQuestionID = CInt(Trim(strQuestionID))
                    strQuestionID = "a" & intQuestionID
                    theAnswerArray = Split(Request.Form.Item(strQuestionID & "_MC"), ",")
                    For k = 0 To UBound(theAnswerArray)
                        intQuestionChoice = Trim(theAnswerArray(k))
                        Call saveEvaluationAnswers(intQuestionID, intQuestionChoice, "")
                        strTextQuestionID = strQuestionID & "text" & intQuestionChoice
                        Call saveTextAnswers(intQuestionID, intQuestionChoice, strTextQuestionID)
                    Next
                ElseIf Left(Request.Form.Keys.Item(x), 1) = "a" Then
                    'These are the radio button choices, just save one choice per question
                    strQuestionID = Request.Form.Keys.Item(x)
                    'strQuestionID = Replace(strQuestionID, "_MC", "")
                    strQuestionID = Replace(strQuestionID, "a", "")
                    intQuestionID = CInt(Trim(strQuestionID))
                    strQuestionID = "a" & intQuestionID
                    intQuestionChoice = Request.Form.Item(strQuestionID)
                    If intQuestionID = 32 Then 'budgetDropDown
                        intQuestionChoice = 190
                    End If
                    Call saveEvaluationAnswers(intQuestionID, intQuestionChoice, "")
                    If intQuestionID = 32 Then 'budgetDropDown
                        intQuestionChoice = Request.Form.Item(x)
                        saveDropDown(intQuestionID, intQuestionChoice)
                    Else
                        strTextQuestionID = strQuestionID & "text" & intQuestionChoice
                        Call saveTextAnswers(intQuestionID, intQuestionChoice, strTextQuestionID)
                    End If

                ElseIf Request.Form.Keys.Item(x) = "EndSurvey" Then
                    Exit For
                End If
                'Then check to see if it's one of the sub-questions
                'Ignore the rest
            End If
        Next

        If Session.Item("iSegmentID") = 1 Then
            Session.Item("bSurvey1") = True
        Else
            Session.Item("bSurvey2") = True
        End If


        'For i = 0 To UBound(arrQuestionArray)
        '	intQuestionID = arrQuestionArray(i)
        '	strQuestionID = "a" & intQuestionID

        '	'Single Answer
        '	If Not IsNothing(Request.Form(strQuestionID)) Then
        '		Call saveEvaluationAnswers(intQuestionID, Request.Form.Item(strQuestionID), "")
        '		strTextQuestionID = strQuestionID & "text" & Request.Form.Item(strQuestionID)
        '		Call saveTextAnswers(intQuestionID, Request.Form.Item(strQuestionID), strTextQuestionID)
        '	ElseIf Not IsNothing(Request.Form(strQuestionID & "_MC")) Then
        '		'MC Multiple Response
        '		theAnswerArray = Split(Request.Form.Item(strQuestionID & "_MC"), ",")
        '		For k = 0 To UBound(theAnswerArray)
        '			intQuestionChoice = Trim(theAnswerArray(k))
        '			Call saveEvaluationAnswers(intQuestionID, intQuestionChoice, "")
        '			strTextQuestionID = strQuestionID & "text" & intQuestionChoice
        '			Call saveTextAnswers(intQuestionID, intQuestionChoice, strTextQuestionID)
        '		Next
        '	End If
        'Next

        ''???what exactly are we doing here...what are we deleting...step through this
        'For i = 1 To UBound(arrTextAnswerArray)
        '	intQuestionID = arrTextAnswerArray(i)
        '	Call deleteAnswers(intQuestionID)
        'Next
    End Sub



	Function GetTestTypeID(ByRef iSegmentID As Integer) As Integer
		Dim TestTypeID As Integer
		Dim objTestTypeID As DataAccess.DANavigation = New DataAccess.DANavigation
		TestTypeID = objTestTypeID.getTestType(iSegmentID)

		Return TestTypeID

	End Function

	Private Sub setSurveyDate()
		'Run the DataUser class and set user classification based on eval answers
		Dim objSetUserClassification As New DataAccess.DAUsers
		objSetUserClassification.SetSurveyDate(iUserAccountID, iSegmentID, iAssessmentID)
	End Sub

	Function getSubQuestion(ByVal subQuestionID) As String

		Dim DRSubQuestion As DataSet
		Dim QuestionTable As DataTable
		Dim QuestionColumn As DataColumn
		Dim QuestionRow As DataRow
		Dim objSubQuestion As New Content.PopulateLists
		Dim strSubQuestions As String
		Dim strQuestions As String
		Dim strTextQuestion As String


		'this just gets the same datareader for comparison to find the subquestions
		DRSubQuestion = objSubQuestion.getQuestions(iSegmentID, iAssessmentID, 1)

		If DRSubQuestion.Tables.Count > 0 Then		  'dataset contains a table...proceed
			For Each QuestionTable In DRSubQuestion.Tables
				For Each QuestionRow In QuestionTable.Rows
					If QuestionRow.Item("iQuestionID") = subQuestionID Then
						strSubQuestions = strSubQuestions & "<input type=text value="""
						strSubQuestions = strSubQuestions & QuestionRow.Item("vchFreeResponse")
						strSubQuestions = strSubQuestions & """ name=a" & subQuestionID & "text" & QuestionRow.Item("iQuestionChoiceID") & " size=20 maxlength=500>"
                        'strSubQuestions = strSubQuestions & "<input type=hidden name=a" & subQuestionID & " value=" & QuestionRow.Item("iQuestionChoiceID") & ">"
						'???the old routine was appending stuff to strings found above...see old routine below
						strQuestions = strQuestions & subQuestionID & ","
						strTextQuestion = strTextQuestion & "," & subQuestionID
					End If
				Next
			Next
		End If

		Return strSubQuestions

	End Function

    Function getSubQuestion(ByVal subQuestionID As Integer, ByVal QuestionTable As DataTable) As String
        Dim QuestionColumn As DataColumn
        Dim strSubQuestions, strTextQuestion, strQuestions, strThisQuestion, strChecked As String
        strThisQuestion = ""
        Dim bHasAnswer As Boolean
        Dim bMakeDiv As Boolean
        Dim QuestionRow As DataRow
        Dim FreeResponse As String
        For Each QuestionRow In QuestionTable.Rows
            strChecked = ""
            If QuestionRow.Item("iQuestionTypeID") = 6 And QuestionRow.Item("iQuestionID") = subQuestionID Then
                strSubQuestions = strSubQuestions & "<input type=text value="""
                strSubQuestions = strSubQuestions & QuestionRow.Item("vchFreeResponse")
                strSubQuestions = strSubQuestions & """ name=a" & subQuestionID & "text" & QuestionRow.Item("iQuestionChoiceID") & " size=20 maxlength=500>"
                'strSubQuestions = strSubQuestions & "<input type=hidden name=a" & subQuestionID & " value=" & QuestionRow.Item("iQuestionChoiceID") & ">"
                strQuestions = strQuestions & subQuestionID & ","
                strTextQuestion = strTextQuestion & "," & subQuestionID
                bMakeDiv = False
            ElseIf QuestionRow.Item("iQuestionTypeID") = 4 And QuestionRow.Item("iQuestionID") = subQuestionID Then
                bMakeDiv = True
                FreeResponse = QuestionRow.Item("vchFreeResponse")
                bHasAnswer = checkForAnswer(QuestionRow.Item("iQuestionChoiceID"), FreeResponse, QuestionRow.Item("iResponse"))
                If bHasAnswer Then
                    strChecked = " checked"
                End If
                If strThisQuestion = QuestionRow.Item("vchQuestionNumber") Then
                    'This isn't the first choice
                    strAllPopUpDivs = strAllPopUpDivs & "<input type=""checkbox"" " & strChecked & " name=""a" & QuestionRow.Item("iQuestionID") & "_MC"" value=""" & QuestionRow.Item("iQuestionChoiceID") & """> " & QuestionRow.Item("vchQuestionChoice") & "<br />"

                Else
                    'This is the first choice
                    strThisQuestion = QuestionRow.Item("vchQuestionNumber")
                    strAllPopUpDivs = strAllPopUpDivs & "<div class=""popupq"" id=""h" & subQuestionID & """ style=""display:none;""><b>" & _
                    QuestionRow.Item("txQuestion") & "</b><br />" & _
                    "<input type=""checkbox"" " & strChecked & " name=""a" & QuestionRow.Item("iQuestionID") & "_MC"" value=""" & QuestionRow.Item("iQuestionChoiceID") & """> " & QuestionRow.Item("vchQuestionChoice") & "<br />"
                End If

                strSubQuestions = ""
            End If
            'If QuestionRow.Item("iQuestionID") = subQuestionID Then
            '    strSubQuestions = strSubQuestions & "<input type=text value="""
            '    strSubQuestions = strSubQuestions & QuestionRow.Item("vchFreeResponse")
            '    strSubQuestions = strSubQuestions & """ name=a" & subQuestionID & "text" & QuestionRow.Item("iQuestionChoiceID") & " size=20 maxlength=500>"
            '    strSubQuestions = strSubQuestions & "<input type=hidden name=a" & subQuestionID & " value=" & QuestionRow.Item("iQuestionChoiceID") & ">"
            '    strQuestions = strQuestions & subQuestionID & ","
            '    strTextQuestion = strTextQuestion & "," & subQuestionID
            'End If
        Next

        If bMakeDiv Then
            strAllPopUpDivs = strAllPopUpDivs & "<br /><br /><a href=""#a"" onclick=""javascript:popup('h" & subQuestionID & "')"">Continue</a>"
            strAllPopUpDivs = strAllPopUpDivs & "</div>"
        End If

        Return strSubQuestions

    End Function

	Private Sub saveEvaluationAnswers(ByVal intQuestionID As Integer, ByVal iQuestionChoiceID As Integer, ByVal strTextAnswer As String)
		'Run the DataAccess class and save interview answers
		Dim objSaveInterviewAnswers As New DataAccess.DAAssessment
		objSaveInterviewAnswers.SaveAnswers(intQuestionID, iAssessmentID, iQuestionChoiceID, strTextAnswer)
	End Sub

	Private Sub deleteAnswers(ByVal intQuestionID As Integer)
		'Run the DataAccess class and delete text interview answers
		Dim objDeleteAnswers As New DataAccess.DAAssessment
		objDeleteAnswers.DeleteAnswers(intQuestionID, iAssessmentID)
	End Sub

	Function GetAccessKey(ByVal intAccessKey As Integer) As String
		Dim strAccessKey As String

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

	Sub saveTextAnswers(ByVal intQuestionID As Integer, ByVal intQuestionChoiceID As Integer, ByVal strTextQuestionID As String)
		'Dim cleanSQL() As String
		Dim strTextAnswer As String
		If Not IsNothing(Request.Form.Item(strTextQuestionID)) Then
			strTextAnswer = Trim(Request.Form.Item(strTextQuestionID))
			If (Not IsNothing(strTextAnswer)) Then			 '1/3/01
				strTextAnswer = cleanSQL(strTextAnswer)
			End If
			Call saveEvaluationAnswers(intQuestionID, intQuestionChoiceID, strTextAnswer)

		End If
    End Sub

    Sub saveDropDown(ByVal intQuestionID As Integer, ByVal intQuestionChoiceID As Integer)
        'Dim cleanSQL() As String
        Dim strTextAnswer As String
        strTextAnswer = getBudgetValue(intQuestionChoiceID)
        Call saveEvaluationAnswers(intQuestionID, 190, strTextAnswer)


    End Sub

    Function getBudgetValue(ByVal iQID As Integer) As String
        Select Case iQID
            Case 1
                Return "2500"
            Case 2
                Return "10000"
            Case 3
                Return "20000"
            Case 4
                Return "30000"
            Case 5
                Return "40000"
            Case 6
                Return "50000"
            Case 7
                Return "60000"
            Case 8
                Return "70000"
            Case 9
                Return "80000"
            Case 10
                Return "90000"
            Case 11
                Return "95000"
            Case Else
                Return "0"
        End Select
    End Function

	Function cleanSQL(ByRef text As String) As String
		If text <> "" Then
			text = Replace(text, "'", "''")
			cleanSQL = text
		Else
			cleanSQL = ""
		End If
	End Function

	Function getIntroText(ByVal iClientID As Integer) As String
		Dim objTextForDispaly As New DataAccess.DAContent
		Dim dr As SqlClient.SqlDataReader = objTextForDispaly.getTextForDisplay(iSegmentID, iClientID)
		If dr.Read() Then
			Return dr.Item("txIntroduction")
		Else
			Return ""
		End If
		dr.Close()
		dr = Nothing
	End Function

	Function classifySOC() As Boolean
		Dim objClassifySOC As DataAccess.DAUpdateUserData = New DataAccess.DAUpdateUserData
		Return objClassifySOC.classifySOC(iUserAccountID)
	End Function


    Function getCustomText(ByVal iTextTypeID As Integer) As String
        'Dim iClientID As Integer
        'iClientID = Session.Item("iClientID")
        'Dim strMsg As String
        'Dim objCustomText As New DataAccess.DAContent
        'strMsg = objCustomText.GetCustomText(iClientID, iTextTypeID)
        If strMsg = Nothing Or IsDBNull(strMsg) Then
            strMsg = ""
        End If
        Return strMsg
    End Function

    Private Sub displayStudentleadership()
        Dim strText As String
        strText = getCustomText(32)

        If strText <> "" Then
            'KSO 6/24/08: replace Resources at with Click on string instead of putting it outside the box
            strText = Replace(strText, "Resources at ", "Click on the links below to view more <b>student leadership opportunities</b> on your campus:<br /><br />")
            'change the alignment to be to the left
            strText = Replace(strText, "align=""center""", "align=""left""")
            'make the box wider
            strText = Replace(strText, "width=""554""", "width=""654""")
            strText = Replace(strText, "width=""560""", "width=""660""")

            Dim custText As Label
            custText = FindControl("custText")

            'strText = "Click on the links below to view more student leadership opportunities on your campus:<br>" & strText
            custText.Text = strText
            custText.Visible = True
        End If
    End Sub

    Private Function hasCustomSE(ByVal iClientID As Integer) As Boolean
        Dim rsTitle As SqlClient.SqlDataReader
        Dim objTitle As New AlcoholEdu.DataAccess.DAAdminSite
        rsTitle = objTitle.getCustomText(iClientID, 68)
        If rsTitle.Read() Then
            Return True
        Else
            Return False
        End If
        rsTitle.Close()
        rsTitle = Nothing
    End Function

    Private Function makeCustomSE(ByVal iClientID As Integer, ByRef strQuestions As String) As String
        Dim strChoices As String
        Dim strOutput As String = ""
        Dim rsSE As SqlClient.SqlDataReader
        Dim objRS As New AlcoholEdu.DataAccess.DAAdminSite
        rsSE = objRS.getCustomText(iClientID, 68)
        If rsSE.Read() Then
            strChoices = rsSE("txNonHTMLText")
        Else
            Return ""
        End If

        strChoices = Replace(strChoices, "^", "")
        Dim arrChoices() As String = Split(strChoices, "<FLD>")
        Dim i As Integer = 0
        Dim iQuestionNo As Integer = 11549 '11538 live
        Dim qc As Integer = 51533 'lowest qc value of placeholder qc, change to 51461 for live site

        For i = 0 To UBound(arrChoices) - 1
            If arrChoices(i) <> "" Then
                strOutput = strOutput & "<tr><td valign=top width=25 class=""answerNumber1"" style=""answerNumber1"">"
                strOutput = strOutput & "<input onClick=""checkSE();"" onKeyPress=""checkSE();""  type=checkbox  name=""a" & iQuestionNo.ToString() & "_MC"" value=""" & qc.ToString() & """ id=id" & qc.ToString() & "></td><td valign=top class=""answerText"" style=""answerText"">" & arrChoices(i) & "</td></tr>"
            End If
            qc = qc + 1
        Next

        If strOutput <> "" Then
            strQuestions = strQuestions & iQuestionNo.ToString() & ","
        End If

        Return strOutput
    End Function


    Private Function getTestInformation(ByVal iSegmentID As Integer, ByVal iUserAccountID As Integer, ByRef iAssessmentID As Integer, ByRef NumberPerPage As Integer, ByRef strSegmentName As String, ByRef RandomQuestions As Boolean, ByRef iSectionTypeID As Integer, ByRef iTestFormID As Integer) As Boolean
        Dim bGotInfo As Boolean = False
        Dim rs As SqlClient.SqlDataReader
        Dim objDataAccess As New AlcoholEdu.DataAccess.DANavigation
        rs = objDataAccess.getTestInformation(iUserAccountID, iSegmentID)

        Try
            If rs.Read Then
                'FormFlag = CBool(rs("bForm"))			 'This has been removed
                If Not IsDBNull(rs("iAssessmentID")) Then
                    iAssessmentID = CInt(rs("iAssessmentID"))
                Else
                    iAssessmentID = 0
                End If
                NumberPerPage = CInt(rs("iNumberPerPage"))
                strSegmentName = CStr(rs("vchSegmentName"))
                RandomQuestions = CBool(rs("tiRandomQuestion"))
                iSectionTypeID = CInt(rs("iSectionTypeID"))
                iTestFormID = CInt(rs("iTestFormID"))
                bGotInfo = True
            End If
        Catch ex As Exception
            Throw New System.Exception(ex.Message, _
            ex.InnerException)
        Finally
            rs.Close()
        End Try
        Return bGotInfo
    End Function

    Private Sub setTestInformation(ByVal iUserAccountID As Integer, ByVal iSegmentID As Integer)
        Dim iAssessmentID, iNum, iSectionTypeID, iTestFormID As Integer
        Dim strSegName As String
        Dim bRandomQuestions As Boolean
        Dim rs As Boolean
        rs = getTestInformation(iSegmentID, iUserAccountID, iAssessmentID, iNum, strSegName, bRandomQuestions, iSectionTypeID, iTestFormID)
        Session.Item("iAssessmentID") = iAssessmentID
        Session.Item("strSegmentName") = strSegName
        Session.Item("iTestID") = iSectionTypeID
        Session.Item("iTestFormID") = iTestFormID
        ' Response.End()
    End Sub

    Private Function makeArray(ByVal strArrayName As String)
        Dim strArray() As String

        If Trim(strArrayName) = "budget" Then
            strArray = New String() {"Choose a Value", "$0 – $4,999", "$5,000 - $14,999", "$15,000 - $24,999", "$25,000 - $34,999", "$35,000 - $44,999", "$45,000 - $54,999", "$55,000 - $64,999", "$65,000 - $74,999", "$75,000 - $84,999", "$85,000 - $94,999", "$95,000+"}
        Else 'Trim(strArrayName) = "error" Then
            strArray = New String() {"Error retrieving values"}
        End If

        Return strArray

    End Function

    Private Function isValidSurvey(ByVal iSegmentID As Integer) As Boolean
        Dim bValid As Boolean = True
        If iSegmentID = 1 Then
            bValid = isSurvey1Valid()
        Else
            bValid = isSurvey2Valid()
        End If
        ' strMsg = strMsg & "<br>TEST"
        Return bValid
    End Function

    Private Function isSurvey1Valid() As Boolean
        'Check each form value against the rules and set any errors in the strMsg variable.
        Dim bSurvey1Valid As Boolean = True
        Dim bQuestionSkipped As Boolean = False 'Need to skip "other" question
        'Loop through the form items that are required questions
        Dim strQuestionID As String
        Dim i As Integer
        For i = 2 To 55
            strQuestionID = Request.Form.Keys.Item(i)
            If Request.Form.Keys.Item(i) = "EndSurvey" Then Exit For
            If InStr(strQuestionID, "a1text") Then
                If Not hasAnswer(i) Then
                    bQuestionSkipped = True
                    bSurvey1Valid = False
                ElseIf Not isNumber(i) Then
                    strMsg = strMsg & "<br>For the first question, please enter a numeric value between 0 and 100,000. If you are having trouble answering the question, please provide your best estimate."
                    bSurvey1Valid = False
                Else
                    'Check value
                    If Not valueValid(100001, 0, i) Then
                        strMsg = strMsg & "<br>For the first question, please enter a numeric value between 0 and 100,000. If you are having trouble answering the question, please provide your best estimate."
                        bSurvey1Valid = False
                    End If
                End If
            ElseIf strQuestionID = "a2_MC" Then
                'Check for a selected answer
                If Not hasAnswer(i) Then
                    bQuestionSkipped = True
                    bSurvey1Valid = False
                End If

                'Dim strQuestionID As String
                'Dim intQuestionID, intQuestionChoice As Integer
                'Dim theAnswerArray() As String
                'strQuestionID = Request.Form.Keys.Item(i)
                'Dim iLen As Integer
                'iLen = Len(Request.Form.Item(strQuestionID))
                'strQuestionID = Replace(strQuestionID, "_MC", "")
                'strQuestionID = Replace(strQuestionID, "a", "")
                'intQuestionID = CInt(Trim(strQuestionID))
                'strQuestionID = "a" & intQuestionID
                'theAnswerArray = Split(Request.Form.Item(strQuestionID & "_MC"), ",")

            Else
                'Exit For
            End If

        Next

        If bSurvey1Valid Then
            If IsNothing(Request.Form("a2_MC")) Then
                bSurvey1Valid = False
                bQuestionSkipped = True
            End If
        End If

        If bQuestionSkipped Then
            strMsg = strMsg & "<br>You have skipped a question that is required for you to receive customized feedback. If you are having trouble answering a question, please provide your best estimate. "
        End If
        If Not bSurvey1Valid Then
            rePopSurvey(1)
        End If
        Return bSurvey1Valid
    End Function

    Private Sub rePopSurvey(ByVal iSurvey As Integer)
        'First get all the responses
        Dim strHTMLQuestions As String
        Dim myControl As New HtmlGenericControl
        iSegmentID = Session.Item("iSegmentID")
        Dim iClientID As Integer
        Dim contentPlaceHolder As New ContentPlaceHolder

        strHTMLQuestions = getEvaluationQuestions(True, iClientID, iSurvey)
        'output questions for display    
        'Dim objInterviewQuestions As HtmlGenericControl
        'objInterviewQuestions = FindControl("evaluationQuestions")
        'objInterviewQuestions.InnerHtml = strHTMLQuestions
        'Dim contentPlaceHolder As New ContentPlaceHolder
        contentPlaceHolder = Master.FindControl("mainContent")
        'Dim myControl As HtmlGenericControl

        strHTMLQuestions = strAllPopUpDivs & strHTMLQuestions
        myControl = contentPlaceHolder.FindControl("evaluationQuestions")
        myControl.EnableViewState = False
        myControl.InnerHtml = strHTMLQuestions
        myControl.DataBind()
    End Sub

    Private Function isSurvey2Valid() As Boolean
        Dim bSurvey2Valid As Boolean = True
        Dim bQuestionSkipped As Boolean = False 'Need to skip "other" question
        'Loop through the form items that are required questions
        Dim i As Integer
        Dim strQuestionID As String
        Dim bSkipTo8 As Boolean
        Dim bSkipToEnd As Boolean
        'Dim arrReqQuestions() As String = {"a31", "a32", "a33", "a34", "a35", "a36", "a37", "a39", "a40", "a41", "a42", "a43", "a44", "a45", "a46"}
        Dim strReqList As String = ", a31, a32, a33, a34, a35, a36, a37, a39, a40, a41, a42, a43, a44, a45, a46"


        For i = 2 To 35
            strQuestionID = Request.Form.Keys.Item(i)
            If Request.Form.Keys.Item(i) = "EndSurvey" Then Exit For
            Select Case strQuestionID
                Case "a38", "a47"
                    'Don't do anything, these aren't checked
                Case "a42"
                    If hasAnswer(i) Then
                        strReqList = Replace(strReqList, ", a42", "")
                    End If
                    If getAnswer(i) = "225" Then
                        bSkipToEnd = True
                        strReqList = Replace(strReqList, ", a43, a44, a45, a46", "")
                    End If
                Case "a35"
                    If hasAnswer(i) Then
                        strReqList = Replace(strReqList, ", a35", "")
                    End If
                    If getAnswer(i) = "202" Then
                        bSkipTo8 = True
                        strReqList = Replace(strReqList, ", a36, a37", "")
                    End If
                Case "a36", "a37"
                    If Not bSkipTo8 Then
                        If hasAnswer(i) Then
                            strReqList = Replace(strReqList, ", " & strQuestionID, "")
                        End If
                    End If
                Case "a43", "a44", "a45", "a46"
                    If Not bSkipToEnd Then
                        If hasAnswer(i) Then
                            strReqList = Replace(strReqList, ", " & strQuestionID, "")
                        End If
                    End If
                Case "a31text189"
                    'This is the FTE question
                    If hasAnswer(i) Then
                        strReqList = Replace(strReqList, ", a31", "")
                    End If
                    If Not isNumber(i) Then
                        'It wasn't a number
                        strMsg = strMsg & "<br/>For question # 1, please enter a numeric value between 0 and 20. If you are having trouble answering the question, please provide your best estimate."
                        bSurvey2Valid = False
                    ElseIf Not decimalValid(20, 0, i) Then
                        'It wasn't in the correct range
                        strMsg = strMsg & "<br/>For question # 1, please enter a numeric value between 0 and 20. If you are having trouble answering the question, please provide your best estimate."
                        bSurvey2Valid = False
                    End If
                Case "a32"
                    'This is the drop-down question
                    If hasAnswer(i) And checkforzero(i) Then
                        strReqList = Replace(strReqList, ", a32", "")
                    End If
                Case "a33", "a34", "a39", "a40", "a41"
                    'These are the rest of the questions
                    If hasAnswer(i) Then
                        strReqList = Replace(strReqList, ", " & strQuestionID, "")
                    End If
            End Select
        Next

        If Len(Trim(strReqList)) > 0 Then
            bQuestionSkipped = True
        End If

        If bQuestionSkipped Then
            strMsg = strMsg & "<br>Please answer all questions on this page.  If you are having trouble answering a question, please provide your best estimate. "
            bSurvey2Valid = False
        End If
        If Not bSurvey2Valid Then
            rePopSurvey(2)
        End If
        Return bSurvey2Valid
    End Function

    Private Function getAnswer(ByVal iIndex As Integer) As String
        If Not IsNothing(Request.Form.Item(iIndex)) Then
            Return Request.Form.Item(iIndex).ToString
        Else
            Return ""
        End If
    End Function

    Private Function hasAnswer(ByVal iIndex As Integer) As Boolean
        If IsNothing(Request.Form.Item(iIndex)) Then
            Return False
        ElseIf Trim(Request.Form.Item(iIndex)) = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function checkforzero(ByVal iIndex As Integer) As Boolean
        If Trim(Request.Form.Item(iIndex)) = "0" Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function isNumber(ByVal iIndex As Integer) As Boolean
        If IsNumeric(Trim(Request.Form.Item(iIndex))) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function valueValid(ByVal iMax As Integer, ByVal iMin As Integer, ByVal iIndex As Integer) As Boolean
        Dim intValue As Integer
        Try
            intValue = CInt(Trim(Request.Form.Item(iIndex)))
        Catch ex As Exception
            Return False
        End Try
        If intValue > iMax Then
            Return False
        ElseIf intValue < iMin Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function decimalValid(ByVal imax As Integer, ByVal iMin As Integer, ByVal iIndex As Integer) As Boolean
        Dim intValue As Decimal
        Try
            intValue = CDec(Trim(Request.Form.Item(iIndex)))
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
        If intValue > imax Then
            Return False
        ElseIf intValue < iMin Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function checkForAnswer(ByVal iQuestionChoiceID As Integer, ByRef vchFreeResponse As String, ByVal iResponse As Integer) As Boolean
        If iQuestionChoiceID = iResponse Then
            Return True
        Else
            vchFreeResponse = ""
            Return False

        End If

    End Function
End Class
