Public Class AdminLogin
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (txtUsername.Text.Equals("admin") And txtPassword.Text.Equals("123456")) Then
            Dim adminDash As AdminDashboard = New AdminDashboard()
            adminDash.Show()
            Me.Close()
        Else
            MessageBox.Show("Username or password is not correct.")
        End If
    End Sub
End Class