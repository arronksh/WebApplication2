Public Class loginForm
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub BTNLOGIN_Click(sender As Object, e As EventArgs) Handles BTNLOGIN.Click
        If tbID.Text = "abc" And tbPass.Text = "abc" Then
            Session("user") = tbID.Text
            Response.Redirect("~/WebForm1.aspx")
        Else
            Response.Write("Login failed")
        End If
    End Sub
End Class