Public Partial Class moreinfo
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack() Then
            If Session.Item("scheduled1") = "yes" Then
                hideForm()
            End If
            If IsNothing(Session.Item("iUserAccountID")) Or Session.Item("iUserAccountID") = "0" Then
                backBtn.Visible = False
                consultBtn.Visible = False
            End If
        End If


    End Sub
    Protected Sub saveConsult_Click(ByVal sender As Object, ByVal e As EventArgs) Handles consultBtn.Click
        'If sYes.Checked Then
        Session.Item("scheduled1") = "yes"
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
        strSub = "Request for Consult " & strRep
        strName = Session.Item("NameTitle")
        strSchool = Session.Item("school")
        strEmail = Session.Item("email")
        strBody = strRep & " should contact " & strName & " at " & strSchool & ". They have requested a consultation from the More Info page regarding their diagnostic results on " & Now().ToString & ". <br><br> You can reach them at " & strEmail & "<br><br>"
        SendTheMessage(strTO, strFrom, strSub, strBody)
    End Sub
    Private Sub hideForm()
        'offerConsult.Visible = False
        'Show the image with the PE name based on the state 
        Dim strImage As String
        strImage = getPEByState() & "1.png"
        consultBtn.ImageUrl = "/images/" & strImage
        consultBtn.AlternateText = "Thank you"
        consultBtn.Enabled = False

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

    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles backBtn.Click
        Response.Redirect("feedback.aspx")
    End Sub
End Class