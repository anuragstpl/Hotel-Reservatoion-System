Public Class AdminDashboard

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim adminDash As ViewCustomers = New ViewCustomers()
        adminDash.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim adminDash As ViewBookings = New ViewBookings()
        adminDash.Show()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim adminDash As AdminLogin = New AdminLogin()
        adminDash.Show()
        Me.Close()
    End Sub
End Class