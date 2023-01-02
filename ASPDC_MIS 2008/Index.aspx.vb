Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlDataReader
Imports System.Web.UI.Page
Imports System.IO

Partial Class Index
    Inherits System.Web.UI.Page
    Dim CS As String = ConfigurationManager.AppSettings("connstring")

    Dim conn As New SqlConnection(CS)
    Dim Cmd As New SqlCommand
    Dim Cmd1 As New SqlCommand
    Dim Cmd2 As New SqlCommand
    Dim Cmd3 As New SqlCommand

    Protected Sub btnLogin_Click(sender As Object, e As System.EventArgs) Handles btnLogin.Click
        If conn.State = ConnectionState.Closed Then conn.Open()
        Dim da As New SqlCommand("SP_ASPDC_UserLogin", conn)
        da.Parameters.Add(New SqlParameter("@LoginName", SqlDbType.VarChar, 50)).Value = txtLoginNm.Value
        da.Parameters.Add(New SqlParameter("@Password", SqlDbType.VarChar, 50)).Value = txtPassword.Value

        da.CommandType = CommandType.StoredProcedure
        Dim reader As SqlDataReader
        reader = da.ExecuteReader

        If reader.Read Then
            Response.Cookies("UserName").Value = reader("PDCUserName")
            Response.Cookies("UserCode").Value = reader("PDCUserId")

            Response.Redirect("UserDashboard.aspx")
            reader.Close()
        Else

            DisplayClientError("Invalid Login Name or Password.")
        End If
    End Sub

    Private Sub DisplayClientError(ByVal errorDesc As String)
        Dim script As String = "alert('" + errorDesc + "');"
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "UserSecurity", script, True)
    End Sub
End Class
