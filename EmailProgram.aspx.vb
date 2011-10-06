Public Partial Class EmailProgram
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack() Then
            populatePage()
        End If
    End Sub

    Private Sub populatePage()
        'Dim iClientID As Integer
        'iClientID = Session.Item("iClientID")

        toEmail.Focus()

        makeParentURL("")
        setUsername()

    End Sub
    Private Sub setUserName()
        Dim strUsername As String
        strUsername = Session.Item("email")
        If InStr(strUsername, "@") Then
            fromEmail.Text = strUsername
            fromEmail.BackColor = Drawing.Color.White
        End If
    End Sub
    Private Function getParentID(ByVal iClientID As Integer) As String
        Dim strID As String
        'strID = getID(iClientID)

        'If strID = "" Then
        strID = "defaultID"
        'End If
        Return strID
    End Function

    Private Sub makeParentURL(ByVal strParentID As String)
        Dim strURL As String
        strURL = "http://www.campusdiagnostic.com"
        parentURL.NavigateUrl = strURL
        parentURL.Text = strURL
    End Sub
    Private Function getID(ByVal iClientID As Integer) As String
        Dim objGetID As New AlcoholEdu.DataAccess.DAClient
        Return objGetID.getParentCourseID(iClientID)
    End Function


    Protected Sub btnSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSend.Click
        'Send the email and show thank you message
        If validateForm() Then
            sendEmail()
            showSuccess()
        Else
            showError()
        End If
    End Sub

    Private Sub sendEmail()
        Dim iClientID, iEmailBatchID, iErr, i As Integer
        iClientID = Session.Item("iClientID")
        Dim strTo, strFrom, strSubject, strBody, strURL, strLink, strChamp As String
        strTo = toEmail.Text
        strFrom = fromEmail.Text
        strSubject = subject.Text

        strChamp = "" 'vbCrLf & vbCrLf & "Click Here: http://hs.outsidetheclassroom.com/become-a-champion-for-alcoholedu if you want to support bringing more effective alcohol education to high schools or colleges and universities in your area." & vbCrLf & vbCrLf

        strURL = "http://www.campusdiagnostic.com"
        strLink = vbCrLf & vbCrLf & "Here's the link: " & strURL & vbCrLf & vbCrLf
        strBody = emailBody1.Text & strLink & strChamp & emailBody2.Text

        If InStr(strTo, ",") Then
            Dim strEmailList() As String
            strEmailList = Split(toEmail.Text, ",")
            For i = 0 To strEmailList.Length - 1
                strTo = strEmailList(i)
                iErr = sendTheEmail(strTo, strFrom, strSubject, strBody, iClientID)
            Next
        Else
            iErr = sendTheEmail(strTo, strFrom, strSubject, strBody, iClientID)
        End If
    End Sub
    Private Function sendTheEmail(ByVal strTo, ByVal strFrom, ByVal strSubject, ByVal strBody, ByVal iClientID)
        Dim iEmailBatchID, iErr As Integer
        Try
            Dim objEmail As New AlcoholEdu.Content.SendEmail
            objEmail.sendEmail(strTo, strFrom, strSubject, strBody)
            'objEmail.sendHTMLEmail(strTo, strFrom, strSubject, strBody)
            'Need to log this
            Dim objEmailBatch As New AlcoholEdu.DataAccess.DAEmail
            iEmailBatchID = objEmailBatch.emailBatch(iClientID, 70)
            Dim objEmailLog As New AlcoholEdu.DataAccess.DAEmail
            iErr = objEmailLog.logEmailSent(iEmailBatchID, strTo)
        Catch ex As Exception

        End Try
        Return iErr
    End Function

    Private Function validateForm() As Boolean
        Dim bValid As Boolean = True
        Dim i As Integer
        If InStr(toEmail.Text, ",") Then
            Dim strEmailList() As String
            strEmailList = Split(toEmail.Text, ",")
            While i < strEmailList.Length
                If Not isValidEmail(strEmailList(i)) Then
                    bValid = False
                End If
                i = i + 1
            End While
        Else
            If Not isValidEmail(toEmail.Text) Then
                bValid = False
            ElseIf Not isValidEmail(fromEmail.Text) Then
                bValid = False
            End If
        End If

        Return bValid
    End Function
    Private Sub showError()
        showForm.Visible = True
        errorMsg.Text = "Please make sure To and From fields contain valid email addresses."
        errorMsg.Visible = True
        errorMsg.Attributes.Add("class", "error")
    End Sub

    Private Sub showSuccess()
        showMessage.Visible = True
        showForm.Visible = False
    End Sub


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

End Class