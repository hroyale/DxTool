Public Partial Class message
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
  
    End Sub


    Protected Sub BtnStart_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStartImg.Click
        Dim iSurvey As Integer

        If Not IsNothing(Session.Item("iSegmentID")) Then
            If Session.Item("iSegmentID") = 1 Or Session.Item("iSegmentID") = 2 Then
                iSurvey = Session.Item("iSegmentID")
            End If
        End If

        If Not IsNothing(Session.Item("iUseraccountID")) And Session.Item("iUserAccountID") <> 0 Then
            Response.Redirect("/survey/evaluation.aspx?s=" & iSurvey.ToString)
        Else
            Response.Redirect("default.aspx")
        End If
    End Sub
End Class