Public Partial Class PrinterFriendly
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack() Then
            fillpage()
        End If
    End Sub

    Private Sub fillpage()
        Dim strContent As String
        strContent = Session.Item("strPF")
        results.Text = strContent
    End Sub

End Class