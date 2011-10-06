Imports System.Data
Imports System.Web.UI.HtmlControls


Partial Public Class Feedback
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim bHasBoth As Boolean
        If Not IsPostBack() Then
            If IsNothing(Session.Item("iUserAccountID")) Or Session.Item("iUserAccountID") = "0" Then
                'send back to registration page
                Response.Redirect("/default.aspx")
            End If
            'If Session.Item("iSegmentID") = 2 Then
            '    If Session.Item("bSurvey1") = True Then
            '        bHasBoth = True
            '    End If
            '    ScorePart2(bHasBoth)
            '    surveyName.Text = "Institutionalization Diagnostic"
            '    If Not bHasBoth Then
            '        otherSurvey.Text = "Try Alcohol Programming Diagnostic"
            '    Else
            '        otherSurvey.Visible = False
            '    End If
            'Else
            '    If Session.Item("bSurvey2") = True Then
            '        bHasBoth = True
            '    End If
            '    scorePart1(bHasBoth)
            '    surveyName.Text = "Alcohol Programming Diagnostic"
            '    If bHasBoth Then
            '        otherSurvey.Visible = False
            '    End If
            'End If
            scorePart1(True)
            ScorePart2(True)
            otherSurvey.Visible = True 'Now the link to How the grades were calculated
            'surveyName.Text = "Programming and Institutionalization Diagnostics"
            hideShowForm()
        End If
    End Sub
    Private Sub hideShowForm()
        If Not IsNothing(Session.Item("scheduled")) Then
            hideForm()
        End If
    End Sub

    Sub ScorePart2(ByVal bHasBoth As Boolean)
        Dim iUserAccountID, iTestFormID, i As Integer
        iUserAccountID = Session.Item("iUserAccountID")
        iTestFormID = 2 'This is the first survey 
        Dim DT As DataTable
        Dim strFTE, strbudget As String
        DT = getAnswers(iUserAccountID, iTestFormID)
        Dim iScore As Double
        Dim bSkipto8 As Boolean = False
        Dim bSkiptoEnd As Boolean = False
        If DT.Rows.Count() > 0 Then
            'We have answers Loop through and grab the score
            i = 0
            Do While i < DT.Rows.Count()
                Select Case DT.Rows(i).Item("iQuestionID")
                    Case 31, 32
                        'These are the two free text answers
                        If DT.Rows(i).Item("iQuestionID") = 31 Then
                            strFTE = DT.Rows(i).Item("vchFreeResponse")
                        Else
                            strbudget = DT.Rows(i).Item("vchFreeResponse")
                        End If
                    Case Else
                        If Not IsDBNull(DT.Rows(i).Item("iChoiceScore")) Then
                            If DT.Rows(i).Item("iQuestionID") = 35 Then
                                If DT.Rows(i).Item("iChoiceScore") = 0 Then
                                    bSkipto8 = True
                                End If
                            ElseIf DT.Rows(i).Item("iQuestionID") = 42 Then
                                If DT.Rows(i).Item("iChoiceScore") = 0 Then
                                    bSkiptoEnd = True
                                End If
                            End If
                            If Not skipThis(DT.Rows(i).Item("iQuestionID"), bSkipto8, bSkiptoEnd) Then
                                iScore = iScore + DT.Rows(i).Item("iChoiceScore")
                            End If
                        End If

                End Select
                i = i + 1
            Loop
        End If

        Dim iClassSize As Integer
        Dim iBudget, iBudgetPerStudent, iFTE, iFTEperStudent As Decimal

        If Not IsNothing(Session.Item("classSize")) Then
            iClassSize = Session.Item("classSize")
        End If
        If IsNumeric(Trim(strbudget)) Then
            iBudget = CDec(Trim(strbudget))
        Else
            iBudget = 0.0
        End If
        If Not IsNothing(iClassSize) And Not IsNothing(iBudget) Then
            If iClassSize > 0 And iBudget > 0 Then
                iBudgetPerStudent = iBudget / iClassSize
            End If
        End If
        If IsNumeric(Trim(strFTE)) Then
            iFTE = CDec(Trim(strFTE))
        Else
            iFTE = 0.0
        End If
        If Not IsNothing(iClassSize) And Not IsNothing(iFTE) Then
            If iClassSize > 0 And iFTE > 0 Then
                iFTEperStudent = iFTE / iClassSize
            End If
        End If
        iScore = iScore + getFTEValue(iFTEperStudent)
        iScore = iScore + getBudgetValue(iBudgetPerStudent)
        'notes.Text = "Prevention Budget Per Student: $" & iBudgetPerStudent
        'notes.Visible = True
        ProgramScore.Visible = True
        InstScore.Visible = True
        budget.Text = FormatCurrency(iBudgetPerStudent).ToString()
        Dim strInstGrade As String
        strInstGrade = getInstLetterGrade(iScore)
        theScore.Text = strInstGrade 'iScore.ToString & " Letter Grade:" & strInstGrade & "<br><br><B>Testing the Custom PE by State Functionality. Your PE rep is: " & getPEByState() & "</b>"


        'Right now this gets sent every time the page is hit. 
        Dim bSent As Boolean
        If Not IsNothing(Session.Item("Notified")) Then
            If Session.Item("Notified") = 1 Then
                bSent = True
            End If
        End If

        If Not bSent Then
            Dim strEmailpart2Body As String
            strEmailpart2Body = makePart2Email(strInstGrade, "", "")
            'strEmailpart2Body = "Prevention budget per student: " & budget.Text & "<br>"
            'strEmailpart2Body = strEmailpart2Body & "Institutionalization  Score: " & theScore.Text & "<br>"
            Dim strName, strSchool, strEmail, strState As String

            strName = Session.Item("NameTitle")
            strSchool = Session.Item("school")
            strEmail = Session.Item("email")
            strState = Session.Item("state")
            Dim strHTMLEmail As String
            strHTMLEmail = startHTMLEMail()
            strHTMLEmail = strHTMLEmail & "<div class=""teal"">Diagnostic Results for " & strSchool & " prepared by " & strName & " on " & Today().ToLongDateString & "</div>"
            strHTMLEmail = strHTMLEmail & Session.Item("strEmailPart1") & strEmailpart2Body
            strHTMLEmail = strHTMLEmail & endHTMLEmail(False)


            'This sends the email with their results to the person who took it
            sendEmail(strHTMLEmail)

            'This sends the email with their results to the PD person (Alexa or the person assigned to them??)


            Dim strPDEmail As String
            strPDEmail = startHTMLEMail()
            strPDEmail = strPDEmail & "<div>Participant Info: " & strName & " from " & strSchool & " in " & strState & "<br> Email Address: " & strEmail & "</div>" & vbCrLf
            strPDEmail = strPDEmail & Session.Item("strEmailPDPt1") & Session.Item("strEmailPDPt2")
            strPDEmail = strPDEmail & endHTMLEmail(True)


            sendPDEmail(strPDEmail)
        End If


    End Sub
    Private Function skipThis(ByVal iQuestionID As Integer, ByVal bSkipto8 As Boolean, ByVal bSkiptoEnd As Boolean) As Boolean
        Select Case iQuestionID
            Case 36, 37
                If bSkipto8 Then
                    Return True
                Else
                    Return False
                End If
            Case 43, 44, 45, 46
                If bSkiptoEnd Then
                    Return True
                Else
                    Return False
                End If
            Case Else
                Return False

        End Select
    End Function

    Private Function getFTEValue(ByVal iFTEPerStudent As Decimal) As Double
        Select Case iFTEPerStudent
            Case Is <= 0.00019999
                Return -2
            Case Is >= 0.0006
                Return 3
            Case Is >= 0.0005
                Return 2
            Case Is > 0.0004
                Return 1
            Case 0.0004
                Return 0
            Case Is > 0.00019999
                Return -1
        End Select
    End Function

    Private Function getBudgetValue(ByVal iBudgetPerStudent As Decimal) As Double
        Select Case iBudgetPerStudent
            Case Is <= 6.5
                Return -2
            Case Is >= 15.0
                Return 3
            Case Is >= 13.0
                Return 2
            Case Is > 9.75
                Return 1
            Case Is = 9.75
                Return 0
            Case Is >= 6.5
                Return -1
        End Select
    End Function

    Private Function getInstLetterGrade(ByVal iScore As Decimal) As String
        Select Case iScore
            Case Is <= -1
                ShowInstF.Visible = True
                Return "F"
            Case Is > 32
                ShowInstA2.Visible = True
                Return "A+"
            Case Is > 29
                ShowInstA.Visible = True
                Return "A"
            Case Is > 26
                ShowInstA1.Visible = True
                Return "A-"
            Case Is > 23
                ShowInstB2.Visible = True
                Return "B+"
            Case Is > 20
                ShowInstB.Visible = True
                Return "B"
            Case Is > 17
                ShowInstB1.Visible = True
                Return "B-"
            Case Is > 14
                ShowInstC2.Visible = True
                Return "C+"
            Case Is > 11
                ShowInstC.Visible = True
                Return "C"
            Case Is > 8
                ShowInstC1.Visible = True
                Return "C-"
            Case Is > 5
                ShowInstD2.Visible = True
                Return "D+"
            Case Is > 2
                ShowInstD.Visible = True
                Return "D"
            Case Is > -1
                ShowInstD1.Visible = True
                Return "D-"
            Case Else
                'Should be impossible
                Return iScore.ToString
        End Select
    End Function

    Sub scorePart1(ByVal bHasBoth As Boolean)
        Dim iUserAccountID, iTestFormID As Integer
        iUserAccountID = Session.Item("iUserAccountID")
        iTestFormID = 1 'This is the first survey 
        Dim DT As DataTable
        DT = getAnswers(iUserAccountID, iTestFormID)

        Dim i As Integer = 0
        Dim iQuestionID As Integer
        Dim iQuestionChoiceID As Integer
        Dim vchQuestionChoice, strOther As String
        Dim iBucket, iUniversalCount, iIndicatedCount, iSelectedCount, iClassSize, iBonusCount, iRawN As Integer
        Dim dblImpactScore, iRawImpact, iRawCost As Double
        Dim dblIndicatedImpact, dblUniversalImpact, dblSelectedImpact As Double
        Dim dblTotalCost, dblCostPerImpactPoint As Double
        Dim hashtable As Hashtable = New Hashtable()
        Dim htUniversal As Hashtable = New Hashtable()
        Dim htSelected As Hashtable = New Hashtable()
        Dim htIdicated As Hashtable = New Hashtable()
        Dim iMax As Integer = 5 'This is the max we're counting in each category
        Dim iHighVal As Integer = 5
        Dim iLowVal As Integer = 1
        Dim iMedVal As Integer = 3
        Dim iSimpleScore As Integer
        Dim iHCount, iMCount, ILCount As Integer
        Dim strNotes As String = ""
        Dim strExample As String = ""
        '----------
        Dim dblI, dblS, dblU As Double

        If DT.Rows.Count() > 0 Then
            'We have answers Loop through and grab the score
            Do While i < DT.Rows.Count()
                Select Case DT.Rows(i).Item("iQuestionID")
                    Case 1
                        If Not IsDBNull(DT.Rows(i).Item("vchFreeResponse")) Then
                            If IsNumeric(DT.Rows(i).Item("vchFreeResponse")) Then
                                iClassSize = CInt(DT.Rows(i).Item("vchFreeResponse"))
                                Session.Item("classSize") = iClassSize
                            End If
                        End If
                        If Not IsDBNull(DT.Rows(i).Item("vchComments")) Then
                            If DT.Rows(i).Item("vchComments") <> "" And strNotes = "" Then
                                strNotes = DT.Rows(i).Item("vchComments")
                                strExample = DT.Rows(i).Item("vchQuestionChoice")
                                'strNotes = strNotes & "<br />" & DT.Rows(i).Item("vchComments")
                            End If
                        End If
                    Case 2 'Programs
                        'Loop Through each and calculate the scores. 

                        If DT.Rows(i).Item("iSubQuestionID") > 0 And DT.Rows(i).Item("iSubQuestionID") <> 23 And DT.Rows(i).Item("iSubQuestionID") <> 22 Then
                            'Save for later, need to read the impact when looping through the choices and then add up what's left.
                            If (Not IsDBNull(DT.Rows(i).Item("iChoiceScore"))) And ((IsDBNull(DT.Rows(i).Item("dcUniversal")) = False) Or (IsDBNull(DT.Rows(i).Item("dcSelected")) = False) Or (IsDBNull(DT.Rows(i).Item("dcIndicated")) = False) = True) Then
                                If Not IsDBNull(DT.Rows(i).Item("dcUniversal")) Then
                                    dblU = DT.Rows(i).Item("dcUniversal")
                                Else
                                    dblU = Nothing
                                End If
                                If Not IsDBNull(DT.Rows(i).Item("dcSelected")) Then
                                    dblS = DT.Rows(i).Item("dcSelected")
                                Else
                                    dblS = Nothing
                                End If
                                If Not IsDBNull(DT.Rows(i).Item("dcIndicated")) Then
                                    dblI = DT.Rows(i).Item("dcIndicated")
                                Else
                                    dblI = Nothing
                                End If
                                Dim A1 As New Answers(DT.Rows(i).Item("iSubQuestionID"), DT.Rows(i).Item("iChoiceScore"), dblU, dblS, dblI)
                                hashtable.Add(A1.QuestionID, A1)
                            Else
                                Trace.Write("Error:" & DT.Rows(i).Item("iSubQuestionID").ToString)
                            End If
                        Else
                            'This didn't have choices use default bucket
                            iBucket = DT.Rows(i).Item("iChoiceScore")
                            If iBucket = 1 Then
                                iIndicatedCount = iIndicatedCount + 1
                                If Not IsDBNull(DT.Rows(i).Item("dcIndicated")) Then
                                    dblIndicatedImpact = dblIndicatedImpact + DT.Rows(i).Item("dcIndicated")
                                End If
                            ElseIf iBucket = 2 Then
                                iSelectedCount = iSelectedCount + 1
                                If Not IsDBNull(DT.Rows(i).Item("dcSelected")) Then
                                    dblSelectedImpact = dblSelectedImpact + DT.Rows(i).Item("dcSelected")
                                End If
                            ElseIf iBucket = 3 Then
                                iUniversalCount = iUniversalCount + 1
                                If Not IsDBNull(DT.Rows(i).Item("dcUniversal")) Then
                                    dblUniversalImpact = dblUniversalImpact + DT.Rows(i).Item("dcUniversal")
                                End If
                            End If
                        End If
                        'If Not IsDBNull(DT.Rows(i).Item("dcImpact")) Then
                        '    iRawImpact = iRawImpact + DT.Rows(i).Item("dcImpact")
                        '    '    If IsDBNull(DT.Rows(i).Item("dcImpact")) > 35 Then
                        '    '        'Count as High
                        '    '        iHCount = iHCount + 1
                        '    '        If iHCount <= iMax Then
                        '    '            iSimpleScore = iSimpleScore + iHighVal
                        '    '        End If
                        '    '    ElseIf IsDBNull(DT.Rows(i).Item("dcImpact")) > 10 Then
                        '    '        'count as med
                        '    '        iMCount = iMCount + 1
                        '    '        If iMCount <= iMax Then
                        '    '            iSimpleScore = iSimpleScore + iMedVal
                        '    '        End If
                        '    '    Else
                        '    '        'count as low
                        '    '        ILCount = ILCount + 1
                        '    '        If ILCount <= iMax Then
                        '    '            iSimpleScore = iSimpleScore + iLowVal
                        '    '        End If
                        '    '    End If

                        'End If
                        'If Not IsDBNull(DT.Rows(i).Item("dcCost")) Then
                        '    iRawCost = iRawCost + DT.Rows(i).Item("dcCost")
                        'End If
                        iRawN = iRawN + 1
                        If Not IsDBNull(DT.Rows(i).Item("vchComments")) Then
                            If DT.Rows(i).Item("vchComments") <> "" And strNotes = "" Then
                                strNotes = DT.Rows(i).Item("vchComments")
                                strExample = DT.Rows(i).Item("vchQuestionChoice")
                                'strNotes = strNotes & "<br />" & DT.Rows(i).Item("vchComments")
                            ElseIf DT.Rows(i).Item("vchComments") <> "" Then
                                strExample = strExample & ", " & DT.Rows(i).Item("vchQuestionChoice")
                            End If
                        End If

                    Case 23 'Other response
                        If Not IsDBNull(DT.Rows(i).Item("vchFreeResponse")) Then
                            strOther = DT.Rows(i).Item("vchFreeResponse") 'Do I even need this?
                        End If
                        If Not IsDBNull(DT.Rows(i).Item("vchComments")) Then
                            If DT.Rows(i).Item("vchComments") <> "" And strNotes = "" Then
                                strNotes = DT.Rows(i).Item("vchComments")
                                strExample = DT.Rows(i).Item("vchQuestionChoice")
                                'strNotes = strNotes & "<br />" & DT.Rows(i).Item("vchComments")
                            ElseIf DT.Rows(i).Item("vchComments") <> "" Then
                                strExample = strExample & ", " & DT.Rows(i).Item("vchQuestionChoice")
                            End If

                        End If
                    Case 26, 27, 28, 30 'The yes/no questions at the end. Count the yeses
                        If DT.Rows(i).Item("vchQuestionChoice") = "Yes" Then
                            iBonusCount = iBonusCount + 1
                        End If
                        If Not IsDBNull(DT.Rows(i).Item("vchComments")) Then
                            If DT.Rows(i).Item("vchComments") <> "" And strNotes = "" Then
                                strNotes = DT.Rows(i).Item("vchComments")
                                strExample = DT.Rows(i).Item("vchQuestionChoice")
                                'strNotes = strNotes & "<br />" & DT.Rows(i).Item("vchComments")
                            ElseIf DT.Rows(i).Item("vchComments") <> "" Then
                                strExample = strExample & ", " & DT.Rows(i).Item("vchQuestionChoice")
                            End If
                        End If
                    Case Else 'All the pop-up questions

                        If hashtable.Contains(DT.Rows(i).Item("iQuestionID")) Then
                            'We have the score
                            Dim A2 As Answers
                            A2 = CType(hashtable.Item(DT.Rows(i).Item("iQuestionID")), Answers)
                            iBucket = DT.Rows(i).Item("iChoiceScore")
                            If iBucket = 1 Then
                                If htIdicated.Contains(A2.QuestionID) Then
                                Else
                                    iIndicatedCount = iIndicatedCount + 1
                                    'dblIndicatedImpact = dblIndicatedImpact + A2.dcImpact
                                    'This is now the score, not the impact
                                    dblIndicatedImpact = dblIndicatedImpact + A2.dcIndicated
                                    htIdicated.Add(A2.QuestionID, 1)
                                End If

                            ElseIf iBucket = 2 Then
                                If htSelected.Contains(A2.QuestionID) Then
                                Else
                                    iSelectedCount = iSelectedCount + 1
                                    'dblSelectedImpact = dblSelectedImpact + A2.dcImpact
                                    dblSelectedImpact = dblSelectedImpact + A2.dcSelected
                                    htSelected.Add(A2.QuestionID, 2)
                                End If
                            ElseIf iBucket = 3 Then
                                If htUniversal.Contains(A2.QuestionID) Then
                                Else
                                    iUniversalCount = iUniversalCount + 1
                                    'dblUniversalImpact = dblUniversalImpact + A2.dcImpact
                                    dblUniversalImpact = dblUniversalImpact + A2.dcUniversal
                                    htUniversal.Add(A2.QuestionID, 3)
                                End If
                                'Else
                                'Ignore all the 0's

                            End If
                            A2.ChoiceScore = 0
                        End If
                        If Not IsDBNull(DT.Rows(i).Item("vchComments")) Then
                            If DT.Rows(i).Item("vchComments") <> "" Then
                                If strExample = "" Then
                                    strNotes = DT.Rows(i).Item("vchComments")
                                    strExample = DT.Rows(i).Item("vchQuestionChoice")
                                Else
                                    strExample = strExample & ", " & DT.Rows(i).Item("vchQuestionChoice")
                                End If
                                'strNotes = DT.Rows(i).Item("vchComments")
                                'strNotes = strNotes & "<br />" & DT.Rows(i).Item("vchComments")
                            End If
                        End If
                End Select
                i = i + 1
            Loop
        End If
        Dim A As Answers
        Dim e As IEnumerator, de As DictionaryEntry
        e = hashtable.GetEnumerator
        Do While e.MoveNext
            de = CType(e.Current, DictionaryEntry)
            A = CType(de.Value, Answers)
            If A.ChoiceScore > 0 Then
                'We didn't have a sub-question choice so use the default
                iBucket = A.ChoiceScore
                If iBucket = 1 Then
                    iIndicatedCount = iIndicatedCount + 1
                    dblIndicatedImpact = dblIndicatedImpact + A.dcIndicated
                ElseIf iBucket = 2 Then
                    iSelectedCount = iSelectedCount + 1
                    dblSelectedImpact = dblSelectedImpact + A.dcSelected
                ElseIf iBucket = 3 Then
                    iUniversalCount = iUniversalCount + 1
                    dblUniversalImpact = dblUniversalImpact + A.dcUniversal
                End If
                A.ChoiceScore = 0
            End If
        Loop

        iClassSize = Session.Item("classSize")

        Dim dcPercentU, dcPercentI, dcPercentS As Decimal
        dcPercentU = 0.5
        dcPercentI = 0.2
        dcPercentS = 0.3

        dblUniversalImpact = dblUniversalImpact + iBonusCount

        Dim dblCompositeScore As Double

        dblCompositeScore = (dblUniversalImpact * dcPercentU) + (dblIndicatedImpact * dcPercentI) + (dblSelectedImpact * dcPercentS)


        'Now we should have a total count in each bucket and a total score in each bucket 

        'We also have the raw numbers as well.
        'If dblIndicatedImpact <> 0 And iIndicatedCount > 0 Then
        '    dblIndicatedImpact = dblIndicatedImpact / iIndicatedCount
        'Else
        '    dblIndicatedImpact = 0
        'End If
        'If dblSelectedImpact <> 0 And iSelectedCount > 0 Then
        '    dblSelectedImpact = dblSelectedImpact / iSelectedCount
        'Else
        '    dblSelectedImpact = 0
        'End If
        'If dblUniversalImpact <> 0 And iUniversalCount > 0 Then
        '    dblUniversalImpact = dblUniversalImpact / iUniversalCount
        'Else
        '    dblUniversalImpact = 0
        'End If

        'If iClassSize > 0 Then
        '    dblTotalCost = iRawCost * iClassSize
        'Else
        '    dblTotalCost = iRawCost * 10000
        'End If


        'If dblTotalCost > 0 And iRawImpact <> 0 And iRawN > 0 Then
        '    dblCostPerImpactPoint = dblTotalCost / (iRawImpact + iBonusCount / iRawN)
        'End If
        Dim strWarningItem As String = ""
        Dim strWarning As String = ""
        Dim iNoteCount As Integer = 0

        If iIndicatedCount < 2 Then
            'show warning
            strWarningItem = "<b>Indicated</b> "
            iNoteCount = iNoteCount + 1
        End If

        If iSelectedCount < 2 Then
            'show warning
            If strWarningItem <> "" Then
                strWarningItem = strWarningItem & "and <b>Selective</b> "
            Else
                strWarningItem = "<b>Selective</b> "
            End If
            iNoteCount = iNoteCount + 1
        End If

        If iUniversalCount < 2 Then
            'show warning
            If strWarningItem <> "" Then
                strWarningItem = strWarningItem & "and <b>Universal</b> "
            Else
                strWarningItem = "<b>Universal</b> "
            End If
            iNoteCount = iNoteCount + 1
        End If


        If strWarningItem <> "" Then
            strWarning = strWarning & "<br>You should consider adding more programming elements in the <a href=""calculate.aspx#def"" target=""new"">" & strWarningItem & "</a> area(s). Having one or no programs in any area indicates that you have an imbalanced set of prevention programs.<br /><br />"
        End If

        If iRawN > 14 Then
            strWarning = strWarning & "<br>You have chosen to pursue too many programs, thus undermining your efficiency and overall impact. We suggest that you examine your mix of programs to determine which among them are best targeted to your needs and can achieve the greatest impact. By selecting fewer, more targeted programs, you will achieve the greatest cost-effectiveness for your efforts.<br>"
            iNoteCount = iNoteCount + 1
        End If

        If strWarning <> "" Then
            warnings.Text = strWarning
            warnings.Visible = True
            noteHead.Visible = True
        End If


        'TODO: Has this been replaced by the single note? If so store the name of the item as the note and plug into the paragraph.
        If strExample <> "" Then
            strNotes = Replace(getNote(strExample), "<example>", strExample)
            notes.Text = strNotes
            notes.Visible = True
            noteHead.Visible = True
            iNoteCount = iNoteCount + 1
        End If

        If iNoteCount > 3 Then iNoteCount = 3



        Dim strLetterGrade As String
        strLetterGrade = getLetterGrade(dblCompositeScore, iNoteCount)
        LetterGrade.Text = strLetterGrade


        'dblImpactScore = dblUniversalImpact + dblSelectedImpact + dblIndicatedImpact
        'uAverage.Text = Math.Round(dblUniversalImpact, 2).ToString
        'sAverage.Text = Math.Round(dblSelectedImpact, 2).ToString
        'iAverage.Text = Math.Round(dblIndicatedImpact, 2).ToString
        'cAverage.Text = Math.Round(dblImpactScore, 2).ToString
        'cost1.Text = "$" & Math.Round(dblTotalCost, 2).ToString
        'cost2.Text = "$" & Math.Round(dblCostPerImpactPoint, 2).ToString

        makePart1Email(strLetterGrade, strNotes, strWarning)
        Dim strPDEmail As String = Session.Item("strEmailPDPt1")
        Dim strUGrade, strSGrade, strIGrade As String
        strUGrade = getPDLetterGrade(dblUniversalImpact, "U")
        strSGrade = getPDLetterGrade(dblSelectedImpact, "S")
        strIGrade = getPDLetterGrade(dblIndicatedImpact, "I")
        strPDEmail = "<div>Composite Score: " & dblCompositeScore.ToString() & " &nbsp; Universal: " & dblUniversalImpact.ToString() & "/" & strUGrade & "&nbsp;  Selective:" & dblSelectedImpact.ToString() & "/" & strSGrade & "  &nbsp;  Indicated: " & dblIndicatedImpact.ToString() & "/" & strIGrade & " <br></div>" & strPDEmail
        Session.Item("strEmailPDPt1") = strPDEmail

    End Sub



    Private Function getNote(ByRef strExample As String) As String
        If InStr(strExample, ",") Then
            'More than one
            Dim NewString As String
            NewString = StrReverse(Replace(StrReverse(strExample), StrReverse(", "), StrReverse(" and "), , 1))
            strExample = NewString

            Return "You are spending resources on <b> <example> </b> which have no demonstrated efficacy of behavioral change in the research literature. These programs also lack a sufficient theoretical basis to provide any promise of student behavior change. If behavior change is your stated goal for using these programs, then we suggest that you consult the research literature to select a more effective set of programs that have greater potential for impact.<br>"
        Else
            'just one
            Return "You are spending resources on <b> <example> </b> which has no demonstrated efficacy of behavioral change in the research literature. This program also lacks a sufficient theoretical basis to provide any promise of student behavior change. If behavior change is your stated goal for using this program, then we suggest that you consult the research literature to select a more effective set of programs that have greater potential for impact.<br>"
        End If
    End Function
    Private Function getPDLetterGrade(ByVal dblScore As Double, ByVal strType As String) As String
        If strType = "S" Then
            Select Case dblScore
                Case Is <= -2.0
                    Return "F"
                Case Is >= 10
                    Return "A+"
                Case Is > 8.99
                    Return "A"
                Case Is > 7.99
                    Return "A-"
                Case Is > 6.99
                    Return "B+"
                Case Is > 5.99
                    Return "B"
                Case Is > 4.99
                    Return "B-"
                Case Is > 3.99
                    Return "C+"
                Case Is > 2.99
                    Return "C"
                Case Is > 1.99
                    Return "C-"
                Case Is > 0.99
                    Return "D+"
                Case Is >= 0
                    Return "D"
                Case Is > -1.99
                    Return "D-"
                Case Else
                    'Should be impossible
                    Return dblScore.ToString
            End Select
        Else
            Select Case dblScore
                Case Is <= 2.98
                    Return "F"
                Case Is >= 16
                    Return "A+"
                Case Is > 14.99
                    Return "A"
                Case Is > 13.99
                    Return "A-"
                Case Is > 12.99
                    Return "B+"
                Case Is > 11.99
                    Return "B"
                Case Is > 10.99
                    Return "B-"
                Case Is > 9.99
                    Return "C+"
                Case Is > 8.99
                    Return "C"
                Case Is > 7.99
                    Return "C-"
                Case Is > 6.99
                    Return "D+"
                Case Is > 4.99
                    Return "D"
                Case Is > 2.99
                    Return "D-"
                Case Else
                    'Should be impossible
                    Return dblScore.ToString
            End Select
        End If

    End Function
    Private Function getLetterGrade(ByVal dblScore As Double) As String
        Select Case dblScore
            Case Is < 1.3
                ShowF.Visible = True
                Return "F"
            Case Is > 13.89
                ShowA2.Visible = True
                Return "A+"
            Case Is > 12.89
                ShowA.Visible = True
                Return "A"
            Case Is > 11.89
                ShowA1.Visible = True
                Return "A-"
            Case Is > 10.89
                ShowB2.Visible = True
                Return "B+"
            Case Is > 9.89
                ShowB.Visible = True
                Return "B"
            Case Is > 8.89
                ShowB1.Visible = True
                Return "B-"
            Case Is > 7.89
                ShowC2.Visible = True
                Return "C+"
            Case Is > 6.89
                ShowC.Visible = True
                Return "C"
            Case Is > 5.89
                ShowC1.Visible = True
                Return "C-"
            Case Is > 4.49
                ShowD2.Visible = True
                Return "D+"
            Case Is > 2.69
                ShowD.Visible = True
                Return "D"
            Case Is > 1.29
                ShowD1.Visible = True
                Return "D-"
            Case Else
                'Should be impossible
                Return dblScore.ToString
        End Select
    End Function
    Private Function getLetterGrade(ByVal dblScore As Double, ByVal iNoteCount As Integer) As String
        If dblScore > 14 Then dblScore = 14
        If dblScore > 9.89 And iNoteCount > 0 Then
            If iNoteCount = 3 Then
                dblScore = dblScore - 3
            ElseIf iNoteCount = 2 Then
                dblScore = dblScore - 2
            ElseIf iNoteCount = 1 Then
                dblScore = dblScore - 1
            End If
        End If
        Select Case dblScore
            Case Is < 1.3
                ShowF.Visible = True
                Return "F"
            Case Is > 13.89
                ShowA2.Visible = True
                Return "A+"
            Case Is > 12.89
                ShowA.Visible = True
                Return "A"
            Case Is > 11.89
                ShowA1.Visible = True
                Return "A-"
            Case Is > 10.89
                ShowB2.Visible = True
                Return "B+"
            Case Is > 9.89
                ShowB.Visible = True
                Return "B"
            Case Is > 8.89
                ShowB1.Visible = True
                Return "B-"
            Case Is > 7.89
                ShowC2.Visible = True
                Return "C+"
            Case Is > 6.89
                ShowC.Visible = True
                Return "C"
            Case Is > 5.89
                ShowC1.Visible = True
                Return "C-"
            Case Is > 4.49
                ShowD2.Visible = True
                Return "D+"
            Case Is > 2.69
                ShowD.Visible = True
                Return "D"
            Case Is > 1.29
                ShowD1.Visible = True
                Return "D-"
            Case Else
                'Should be impossible
                Return dblScore.ToString
        End Select
    End Function

    Private Sub makePart1Email(ByVal strLetterScore As String, ByVal strNotes As String, ByVal strWarning As String)
        Dim strEmailPart1Body, strEmailPart1PD As String
        strEmailPart1Body = "<div class=""greenBar"">Programmatic Grade:</div><div id=""scoreCard""><div class=""grade""> " & strLetterScore & "</div>" & vbCrLf

        'strEmailPart1Body = "Programming Grade: " & strLetterScore
        'TODO: for internal email send the actual bucket scores and the composite score
        'TODO: for external email include the grade message

        strEmailPart1Body = strEmailPart1Body & getMessage(strLetterScore, 1)

        If strNotes <> "" Or strWarning <> "" Then
            strEmailPart1Body = strEmailPart1Body & "<span class=""notehead"">Additional feedback about your programming:<br /></span>"
        End If
        If strNotes <> "" Then
            strEmailPart1Body = strEmailPart1Body & vbCrLf & vbCrLf & strNotes
        End If
        If strWarning <> "" Then
            strEmailPart1Body = strEmailPart1Body & vbCrLf & vbCrLf & strWarning
        End If
        'If Not bHasBoth Then
        'Save this for sending later
        strEmailPart1Body = Replace(strEmailPart1Body, "calculate.aspx", "http://www.campusdiagnostic.com/survey/calculate.aspx")
        strEmailPart1Body = strEmailPart1Body & "</div>"
        strEmailPart1PD = strEmailPart1Body 'Add anything else just for PDS (like the question responses??)
        Session.Item("strEmailPart1") = strEmailPart1Body
        Session.Item("strEmailPDPt1") = strEmailPart1PD

    End Sub


    Private Function makePart2Email(ByVal strLetterScore As String, ByVal strNotes As String, ByVal strWarning As String) As String
        Dim strEmailPart2Body, strEmailPart2PD As String
        strEmailPart2Body = "<div class=""greenBar"">Institutionalization  Grade:</div><div id=""scoreCard2""><div class=""grade""> " & strLetterScore & "</div>" & vbCrLf
        'TODO: for internal email send the actual bucket scores and the composite score
        'TODO: for external email include the grade message
        'strEmailPart2PD = strEmailPart2Body
        strEmailPart2Body = strEmailPart2Body & getMessage(strLetterScore, 2)
        Session.Item("part2Feedback") = strEmailPart2Body 'Do I need an extra open div now?
        strEmailPart2Body = strEmailPart2Body & "</div>"
        strEmailPart2PD = strEmailPart2Body
        Session.Item("strEmailPDPt2") = strEmailPart2PD

        Return strEmailPart2Body
    End Function


    Private Function getMessage(ByVal strGrade As String, ByVal iPart As Integer) As String
        If iPart = 2 Then
            Select Case strGrade
                Case "A+"
                    Return IA2.Text
                Case "A"
                    Return IA.Text
                Case "A-"
                    Return IA1.Text
                Case "B+"
                    Return IB2.Text
                Case "B"
                    Return IB.Text
                Case "B-"
                    Return IB1.Text
                Case "C+"
                    Return IC2.Text
                Case "C"
                    Return IC.Text
                Case "C-"
                    Return IC1.Text
                Case "D+"
                    Return ID2.Text
                Case "D"
                    Return ID0.Text
                Case "D-"
                    Return ID1.Text
                Case "F"
                    Return IF0.Text
                Case Else
                    Return ""
            End Select
        Else
            Select Case strGrade
                Case "A+"
                    Return PA2.Text
                Case "A"
                    Return PA.Text
                Case "A-"
                    Return PA1.Text
                Case "B+"
                    Return PB2.Text
                Case "B"
                    Return PB.Text
                Case "B-"
                    Return PB1.Text
                Case "C+"
                    Return PC2.Text
                Case "C"
                    Return PC.Text
                Case "C-"
                    Return PC1.Text
                Case "D+"
                    Return PD2.Text
                Case "D"
                    Return PD.Text
                Case "D-"
                    Return PD1.Text
                Case "F"
                    Return PF.Text
                Case Else
                    Return ""
            End Select
        End If

    End Function
    Private Sub sendEmail(ByVal strBody As String)
        Dim strTO, strFrom, strSub, strName, strSchool, strEmail As String

        strTO = Session.Item("email")
        strFrom = "admin@outsidetheclassroom.com"
        strSub = "Your Diagnostic Results from Outside The Classroom"
        strName = Session.Item("NameTitle")
        strSchool = Session.Item("school")
        strEmail = Session.Item("email")
        'strTO = "haynes@outsidetheclassroom.com;" & strEmail
        SendTheMessage(strTO, strFrom, strSub, strBody)
        'Now send a version to the PSD 
        'strTO = "haynes@outsidetheclassroom.com"
        'strSub = "DX Tool Feedback for " & strSchool
        'If Not IsNothing(Session.Item("scheduled")) Then
        '    If Session.Item("scheduled") = "yes" Then
        '        strBody = "<b>Consult requested. Please contact ASAP.</b><br>" & strBody
        '    Else
        '        strBody = "Said No to consult request.<br>" & strBody
        '    End If
        'End If
        Session.Item("Notified") = 1
    End Sub

    Private Sub sendPDEmail(ByVal strBody As String)
        Dim strTO, strFrom, strSub, strName, strSchool, strEmail As String
        strTO = "dxtool@outsidetheclassroom.com"
        'strTO = Session.Item("email")
        strFrom = "admin@outsidetheclassroom.com"
        strName = Session.Item("NameTitle")
        strSchool = Session.Item("school")
        strEmail = Session.Item("email")
        strSub = "Campus Diagnostic Results for " & strSchool
        SendTheMessage(strTO, strFrom, strSub, strBody)
    End Sub

    Function getAnswers(ByVal iUserAccountID As Integer, ByVal iTestformID As Integer) As DataTable
        Dim objResponses As New AlcoholEdu.DataAccess.DAAssessment
        Return objResponses.getAssessmentResponses(iUserAccountID, iTestformID)
    End Function

    'otherSurvey
    Protected Sub otherSurvey_Click(ByVal sender As Object, ByVal e As EventArgs) Handles otherSurvey.Click
        'Load the other survey
        'Dim iSegmentID As Integer
        'iSegmentID = Session.Item("iSegmentID")
        'If iSegmentID = 1 Then
        '    Session.Item("iSegmentID") = 2
        'Else
        '    Session.Item("iSegmentID") = 1
        'End If

        Response.Redirect("calculate.aspx")

    End Sub

    Protected Sub saveConsult_Click(ByVal sender As Object, ByVal e As EventArgs) Handles consultBtn1.Click
        'If sYes.Checked Then
        Session.Item("scheduled") = "yes"
        hideForm()
        sendYesEmail()
        'ElseIf sNo.Checked Then
        '    Session.Item("scheduled") = "no"
        '    hideForm()
        'Else
        '    'Show a message that they didn't make a choice
        '    offerMsg.Visible = True
        'End If

        'Note: even though it isn't in the non-existent spec, I'm going to send an internal email
        'as soon as "yes" is submitted for consult. Don't want people to click away and not get it scheduled.

    End Sub
    Private Sub sendYesEmail()
        Dim strTO, strFrom, strSub, strBody, strName, strSchool, strEmail, strRep As String
        strRep = getPEByState()
        strTO = "dxtool@outsidetheclassroom.com"
        'Do we want this email to go to the PE rep or to Alexa to coordinate all of them?
        strFrom = "admin@outsidetheclassroom.com"
        strSub = "Request for Consult for " & strRep
        strName = Session.Item("NameTitle")
        strSchool = Session.Item("school")
        strEmail = Session.Item("email")
        strBody = strRep & " should contact " & strName & " at " & strSchool & ". They have requested a consultation on their diagnostic results from the results page on " & Now().ToString & ". <br><br> You can reach them at " & strEmail & "<br><br>"
        SendTheMessage(strTO, strFrom, strSub, strBody)
    End Sub
    Private Sub hideForm()
        'offerConsult.Visible = False
        'Show the image with the PE name based on the state 
        Dim strImage As String
        strImage = getPEByState() & "2.png"
        consultBtn1.ImageUrl = "/images/" & strImage
        consultBtn1.AlternateText = "Thank you"
        consultBtn1.Enabled = False

    End Sub

    Private Sub SendTheMessage(ByVal strTo As String, ByVal strFrom As String, ByVal strSub As String, ByVal strBody As String)
        Dim objSendMail As New AlcoholEdu.Content.SendEmail
        objSendMail.sendHTMLEmail(strTo, strFrom, strSub, strBody)
    End Sub

    Private Function getPEByState() As String
        Dim strState As String = ""
        'strState = Session.Item("state")
        If Not IsNothing(Session.Item("state")) Then
            strState = Session.Item("state")
        End If

        Select Case strState
            Case "AR", "DE", "HI", "LA", "MD", "NJ", "OK", "PA", "TX", "VA", "DC", "AZ", "NM"
                Return "doug"
            Case "CO", "IA", "ID", "IL", "IN", "KS", "KY", "MI", "MN", "MO", "MT", "ND", "NE", "OH", "SD", "TN", "UT", "WI", "WV", "WY", "OR", "WA"
                Return "jerry"
            Case "AL", "CT", "FL", "GA", "MA", "ME", "MS", "NC", "NH", "NY", "RI", "SC", "VT"
                Return "kirby"
            Case Else
                Return "sean"
        End Select
    End Function

    Protected Sub printThisBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles printThisBtn.Click
        createPrinterFriendly()
        'Response.Redirect("printerFriendly.aspx")
        openPrinterPage.Visible = True
    End Sub
    Private Sub createPrinterFriendly()
        Dim strPF, strName, strSchool As String
        strPF = Session.Item("strEmailPart1")
        strPF = strPF & Session.Item("part2Feedback") & "<p></p>"
        strPF = strPF & getCalculate()
        strPF = strPF & getMoreInfo()
        strName = Session.Item("NameTitle")
        strSchool = Session.Item("school")
        strPF = "<div class=""teal"">Diagnostic Results for " & strSchool & " prepared by " & strName & " on " & Today().ToLongDateString & "</div>" & strPF

        Session.Item("strPF") = strPF
    End Sub

    Protected Sub readMoreBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles readMoreBtn.Click
        Response.Redirect("moreinfo.aspx")
    End Sub

    Private Function startHTMLEMail() As String
        Dim strHTML As String
        strHTML = "<HTML><HEAD><TITLE></TITLE><link href=""http://www.campusdiagnostic.com/documentStyleEmail.css"" rel=""stylesheet"" type=""text/css"" /></HEAD><BODY>" & vbCrLf
        strHTML = strHTML & "<div class=""container""><img src=""http://www.campusdiagnostic.com/images/feedback.png"" alt=""The Campus Diagnostic"" style=""border-width:0px;"" />"

        Return strHTML
    End Function

    Private Function endHTMLEmail(ByVal bPD As Boolean) As String
        Dim strContent, strHTML As String
        If bPD Then
            strContent = ""
        Else
            strContent = "<br><a href=""http://hs.outsidetheclassroom.com/schedule-a-consultation"">Schedule a Consultation to Learn More About Your Campus Diagnostic Grade</a><br><div class=""teal"">Read More About: </div> <a href=""http://www.campusdiagnostic.com/survey/calculate.aspx"">How we calculated your grades</A><br>"
            strContent = strContent & "<a href=""http://www.campusdiagnostic.com/survey/moreinfo.aspx"">The <i>full</i> Diagnostic Inventory</A><br>"
            strContent = strContent & "<a href=""http://www.outsidetheclassroom.com/solutions/higher-education/Coalition.aspx"">Outside The Classroom’s Alcohol Prevention Coalition</A><br>"
            strContent = strContent & "<div class=""teal"">Recommended Resources:</div><a href=""http://hs.outsidetheclassroom.com/download-pdf-institutionalizing-alcohol-prevention/"">Institutionalizing Alcohol Prevention: Strategies for Engaging Key Stakeholders and Aligning Objectives (pdf)</a><br>"
            strContent = strContent & "<a href=""http://hs.outsidetheclassroom.com/download-pdf-alcohol-free-options/"">Using Alcohol-Free Options to Promote a Healthy Campus Environment (pdf)</a><br>"
            strContent = strContent & "<a href=""http://www.outsidetheclassroom.com/community/tools-resources/campus-diagnostic/impact-stories.aspx"">How other campuses have used their Campus Diagnostic results to garner support for effective alcohol prevention</A><br>"

            'strContent = getCalculate()
            'strContent = strContent & getMoreInfo()
        End If

        strHTML = "</div></BODY></HEAD>"
        Return strContent & vbCrLf & strHTML
    End Function

    Private Function getCalculate() As String
        Dim strContent As String

        strContent = "</div><div class=""infopage""><div class=""teal"">How we calculated your grades</div><div class=""subTeal"">Understanding the Basis of Your Programming Grade</div>The letter grade and qualitative feedback regarding where your institution stands with respect to alcohol prevention programs is based upon our examination and consideration of several key factors, including: <ul> <li>the extent of your prevention programming </li> <li>the strength of your programming based on research evidence, and </li> <li>how the programs you u•	how the programs that you identified are being used to target different segments of your student population.  </li> </ul>" & vbCrLf
        strContent = strContent & "<p>To develop the underlying scoring rubric, our researchers examined more than 200 studies on the relative efficacy of a variety of campus alcohol prevention programs, applying a standardized procedure to assign numeric values to the reported outcomes of these studies, and then averaging the values across the set of studies for each of the programs examined. We also noted the characteristics of the study sample—whether these were a random sample of students, high-risk students, students mandated to receive the program, or other subsets of the student population. For several programs that lacked any evidence of effectiveness in the research literature, we took into consideration whether there was a sound theoretical basis underlying these approaches.  </p>" & vbCrLf
        strContent = strContent & "<p>In developing the score, we consider how campuses target their alcohol prevention efforts, whether they be <ul> <li> <i>universal </i>in nature (targeting the entire student body)</li>  <li> <i>selective </i>(targeting known high-risk student groups), or </li> <li> <i>indicated</i> (targeting students at the early stages of developing alcohol problems)[1].</li>  </ul></p>" & vbCrLf
        strContent = strContent & "<p>Applying a public health model grounded in the <i>prevention paradox[2]</i>, we place greater emphasis on the scores of universal programs versus selective or indicated. </p> "
        strContent = strContent & "<p>In examining the mix of programming a campus may have in place, we do not assume that <em>more is more.</em> As such, we may determine that a campus has <i>too many</i> programming elements,in effect diminishing the impact of their more effective programs by depleting the resources and attention of prevention personnel. </p>"
        strContent = strContent & "<p>In our analysis, we assume optimal fidelity among programs administered by respondents, yet we acknowledge that there can be varying degrees of reliability in the actual implementation of campus alcohol prevention programs. Certain prevention programs whose effectiveness rely heavily on principles of best practice but can also be more susceptible to implementation shortfall include efforts such as BASICS, alcohol-free programming, social norms marketing, and many others. We encourage campuses to consult the many free resources that are available to assist them in strengthening the  implementation of these types of programs. For example, many campuses have found our publication  <a href=""http://hs.outsidetheclassroom.com/download-pdf-alcohol-free-options/"" target=""_new""><i>Using Alcohol-Free Options to Promote a Healthy Campus Environment</i></a> very helpful in creating effective, high-impact alcohol-free programs and events.  </p>" & vbCrLf
        strContent = strContent & "<div class=""subTeal"">Understanding the Basis of Your Institutionalization Grade</div> <p>The feedback we offer on your responses to this part of the campus diagnostic are based on the best practices of institutions that have made breakthrough progress in alcohol prevention, and upon hundreds of interviews with experts,  alcohol prevention professionals and other officials on campuses across the country. </p>" & vbCrLf

        Return strContent

    End Function

    Private Function getMoreInfo() As String
        Dim strContent As String
        strContent = "<div class=""infopage""><div class=""teal"">More information on the full Diagnostic Inventory:</div><div>" & vbCrLf
        strContent = strContent & "The Campus Diagnostic  is comprised of a sample of questions excerpted from two sections of a much more extensive instrument, The Alcohol Prevention Coalition Diagnostic Inventory. The longer tool was developed by Outside The Classroom to comprehensively assesses many dimensions of campus alcohol prevention at  our <a href=""http://www.outsidetheclassroom.com/solutions/higher-education/Coalition.aspx"">Alcohol Prevention Coalition</a> partner institutions. These dimensions extend beyond a close examination of prevention programming and the degree of institutional support for alcohol prevention on campus, scrutinizing campus alcohol policies, their enforcement and adjudication, adherence to processes deemed critical to success in alcohol prevention, and the extent of relationships with a variety of constituencies that are important to prevention success. Completion of this full Diagnostic Inventory allows Alcohol Prevention Coalition campuses to not only pinpoint areas of strength and weakness and set goals for improvement, but also to benchmark their alcohol prevention progress annually and see how they compare to other campuses that have taken the instrument. </div>" & vbCrLf
        strContent = strContent & "<div id=""footNotes""><div id=""ftn1""><p class=""FootnoteText"">[1]In a 1994 report, the Institute of Medicine proposed a framework for classifying prevention based on Gordon's (1987) operational classification of disease prevention. The IOM model divides the continuum of services into three parts: prevention, treatment, and maintenance. The prevention category is divided into three classifications--<i>universal</i>, <i>selective</i>, and <i>indicated</i> prevention. For more information, visit <a href=""http://www.mypreventioncommunity.org/resource/collection/8CC9C598-EF77-4CDB-A2DF-88AB150A4832/25EIOMModel.pdf"" target=""_new"">http://www.mypreventioncommunity.org/resource/collection/8CC9C598-EF77-4CDB-A2DF-88AB150A4832/25EIOMModel.pdf</a> </p></div>" & vbCrLf
        strContent = strContent & "<div id=""ftn2""><p class=""FootnoteText"">[2] The prevention paradox describes a somewhat counterintuitive public health phenomenon, where the greatest negative impact from a disease or disorder occurs among those considered at low or moderate risk of the disease or disorder, and a relatively small degree of negative impact comes from the highest risk population. For more information on how this phenomenon relates to college alcohol prevention efforts, read Weitzman and Nelson’s paper at <a href=""http://www.hsph.harvard.edu/cas/Documents/paradox/Prev_Paradox.pdf"" target=""_new"">http://www.hsph.harvard.edu/cas/Documents/paradox/Prev_Paradox.pdf</a>. </p></div></div>" & vbCrLf

        Return strContent
    End Function
End Class

Public Class Answers
    Public QuestionID As Integer
    Public ChoiceScore As Integer
    Public dcUniversal As Double
    Public dcSelected As Double
    Public dcIndicated As Double

    Public Sub New( _
       ByVal m_QuestionID As Integer, _
       ByVal m_ChoiceScore As Integer, _
       ByVal m_dcUniversal As Double, _
       ByVal m_dcSelected As Double, _
       ByVal m_dcIndicated As Double)
        QuestionID = m_QuestionID
        ChoiceScore = m_ChoiceScore
        dcUniversal = m_dcUniversal
        dcSelected = m_dcSelected
        dcIndicated = m_dcIndicated
    End Sub





End Class