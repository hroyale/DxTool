Public Partial Class calculate
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsNothing(Session.Item("iUserAccountID")) Or Session.Item("iUserAccountID") = "0" Then
            backBtn.Visible = False
        End If
    End Sub
    Protected Sub backBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles backBtn.Click
        Response.Redirect("feedback.aspx")
    End Sub
End Class