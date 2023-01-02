
Partial Class MasterPage
    Inherits System.Web.UI.MasterPage


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        lblUserName.Text = ""
        If Request.Cookies("UserName").Value <> "" Then
            lblUserName.Text = Request.Cookies("UserName").Value
        End If
    End Sub
End Class

