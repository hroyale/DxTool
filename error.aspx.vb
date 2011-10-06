Public Partial Class _error
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Write("TestID " & Session.Item("iTestID"))
    End Sub

End Class