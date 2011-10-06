Partial Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session.Item("iUseraccountID") = 1
        'Session.Item("iAccountID") = 1
        ''Session.Item("iAssessmentID") = 1
        'Session.Item("iSegmentID") = 2
        ''Session.Item("strSegmentName") = "dx survey"
        'Session.Item("iTestID") = 22
        Dim strSurvey = Request.QueryString("schoice")
        If Not IsPostBack Then
            Session.RemoveAll()
            If IsNumeric(strSurvey) Then
                Session.Item("iSegmentID") = CInt(strSurvey)
            Else
                Session.Item("iSegmentID") = 1
            End If
        End If


    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click, btnSubmitImg.Click
        'Create the Account
        loadAccount()

    End Sub

    Protected Function createAccount(ByVal strFName As String, ByVal strLName As String, ByVal strEmail As String, ByVal strTitle As String, ByVal strInstitution As String, ByVal strState As String, ByVal strClassSize As String) As Integer
        Dim objInsert As New AlcoholEdu.DataAccess.DAUsers
        Return objInsert.InsertAccount(strFName, strLName, strEmail, strTitle, strInstitution, strState, strClassSize)
    End Function

    Protected Sub loadAccount()
        Dim strFName, strLName, strEmail, strTitle, strInstitution, strState, strClassSize, strError As String
        Dim iAccountID As Integer
        Dim bNotValid As Boolean = False
        strFName = fName.Text
        strLName = lName.Text
        strEmail = email.Text
        strTitle = title.Text
        strInstitution = school.Text
        strState = State.SelectedValue
        'strClassSize = studentPop.Text
        strClassSize = "0"

        'strError = "Please correct the following errors and resubmit:<br>"
        strError = ""

        'If Not IsNumeric(Trim(strClassSize)) Then
        '    iClassSize = makeNumber(strClassSize)
        '    If iClassSize > 0 Then
        '        strClassSize = iClassSize.ToString
        '        Session.Item("classSize") = iClassSize
        '    Else
        '        strClassSize = ""
        '        strError = strError & "&nbsp; Enter a valid number for Number of Undergraduate Students.<br>"
        '        bNotValid = True
        '    End If
        'Else

        '    Session.Item("classSize") = CInt(Trim(strClassSize))
        'End If
        If strEmail = "" Or strInstitution = "" Or strState = "oo" Or strFName = "" Or strLName = "" Or strTitle = "" Then
            'Send error back to front end
            strError = strError & "<br><br>Please complete the missing information."
            bNotValid = True

        End If
        If Not isValidEmail(strEmail) Then
            strError = strError & "<br><br>Please enter a valid email address."
            bNotValid = True
        End If

        If bNotValid Then
            errorMsg.Text = strError
            errorMsg.Visible = True
        Else
            iAccountID = createAccount(strFName, strLName, strEmail, strTitle, strInstitution, strState, strClassSize)
            Session.Item("iUseraccountID") = iAccountID
            Session.Item("iAccountID") = iAccountID
            'Session.Item("iAssessmentID") = 1
            Session.Item("NameTitle") = strFName & " " & strLName & ", " & strTitle
            Session.Item("School") = strInstitution
            Session.Item("email") = strEmail
            Session.Item("state") = strState
            Dim iSurvey As Integer
            iSurvey = 1
            If Not IsNothing(Session.Item("iSegmentID")) Then
                If Session.Item("iSegmentID") = 1 Or Session.Item("iSegmentID") = 2 Then
                    iSurvey = Session.Item("iSegmentID")
                End If
            End If
            Session.Item("iSegmentID") = 1
            Session.Item("strSegmentName") = "dx survey"
            Session.Item("iTestID") = 20
            Trace.Write("UserAccount " & Session.Item("iUseraccountID"))
            Trace.Write("TestID " & Session.Item("iTestID"))
            'Response.Redirect("/survey/evaluation.aspx?s=" & iSurvey.ToString)
            send2HubSpot()
            Response.Redirect("message.aspx")
        End If


    End Sub
    Private Sub send2HubSpot()
        Dim strURL, strPost As String
        Try
            strPost = "FirstName=" & Server.UrlEncode(fName.Text) & _
            "&LastName=" & Server.UrlEncode(lName.Text) & _
            "&Email =" & Server.UrlEncode(email.Text) & _
            "&Company=" & Server.UrlEncode(school.Text) & _
            "&JobTitle=" & Server.UrlEncode(title.Text) & _
            "&IPAddress=" & Server.UrlEncode(System.Web.HttpContext.Current.Request.UserHostAddress) & _
            "&UserToken=" & Server.UrlEncode(Request.Cookies("hubspotutk").Value.ToString)
            'set POST url
            strURL = " http://hs.outsidetheclassroom.com/campus-diagnostic?app=leaddirector&FormName=campus-diagnostic "
            'create webrequest object and set headers for POST
            Dim _request As System.Net.HttpWebRequest
            _request = System.Net.WebRequest.Create(strURL)
            _request.Method = "POST"
            _request.ContentType = "application/x-www-form-urlencoded"
            _request.ContentLength = strPost.Length
            _request.KeepAlive = False
            'create streamwriter object and send data
            Dim sw As System.IO.StreamWriter
            sw = New System.IO.StreamWriter(_request.GetRequestStream())
            sw.Write(strPost)
            sw.Close()
        Catch ex As Exception
            'Don't want to throw an error if this doesn't work live
        End Try



    End Sub
    Private Function makeNumber(ByVal strString As String) As Integer
        If strString.Length > 0 Then
            strString = Trim(strString)
            strString = Replace(strString, ",", "")
            strString = Replace(strString, "k", "000")
            strString = Replace(strString, "~", "")
            If IsNumeric(strString) Then
                Return CInt(strString)
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function


    Protected Sub createAssessment()

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

        Return ValidFlag
    End Function
End Class