Public Class ChooseDateTime

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.IsDateChoosen = True
        Reservation.Checkin = DateTimePicker1.Value.ToString
        Reservation.Checkout = DateTimePicker2.Value.ToString
        Me.Hide()
        Form1.confirmBooking1()
    End Sub
End Class