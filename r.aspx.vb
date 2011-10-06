Public Partial Class r
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Status = "302 Moved Temporarily"
        Response.AddHeader("Location", "http://glockanalytics.com/v3/" + Request.Url.Query)

    End Sub

End Class