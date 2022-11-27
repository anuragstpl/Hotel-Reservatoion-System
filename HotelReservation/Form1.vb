Imports System.Data.OleDb
Imports System.Math
Imports System.Net.Mail
Imports System.Threading.Thread

Public Class Form1
    Dim conDatabase As OleDbConnection
    Dim cmdDatabase As OleDbCommand = New OleDbCommand
    Dim pnInfo As Boolean = False
    Dim rxSms, trash, loginMode, loginSuccess, room, email, tempEmail As String
    Dim roomPrice As Integer
    Dim k As Integer = 0
    Dim l As Integer = 8
    Dim Mail As New MailMessage
    Dim SMTP As New SmtpClient("smtp.gmail.com")


    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        websiteDisplay()
    End Sub

    Private Sub websiteDisplay()
        Try
            If Not txtWebAdd.Text = "www.yayahotel.com" And Not txtWebAdd.Text = "yayahotel.com" And Not txtWebAdd.Text = "http://www.yayahotel.com" Then
                pnYaya.Visible = False
                pnYaya.SendToBack()
                WebBrowser1.Visible = True
                WebBrowser1.BringToFront()
                If txtWebAdd.Text = "[Blank Page]" Then

                ElseIf txtWebAdd.Text.Substring(0, 4) = "http" Then
                    WebBrowser1.Navigate(txtWebAdd.Text)
                    txtWebAdd.Text = txtWebAdd.Text
                ElseIf txtWebAdd.Text.Substring(0, 4) = "www." Then
                    WebBrowser1.Navigate("http://" + txtWebAdd.Text)
                    txtWebAdd.Text = "http://" + txtWebAdd.Text
                Else
                    WebBrowser1.Navigate("http://www." + txtWebAdd.Text)
                    txtWebAdd.Text = "http://www." + txtWebAdd.Text
                End If
            Else
                WebBrowser1.Visible = False
                WebBrowser1.SendToBack()
                pnYaya.Visible = True
                pnYaya.BringToFront()
                If txtWebAdd.Text = "[Blank Page]" Then

                ElseIf txtWebAdd.Text.Substring(0, 4) = "http" Then
                    txtWebAdd.Text = txtWebAdd.Text
                ElseIf txtWebAdd.Text.Substring(0, 4) = "www." Then
                    txtWebAdd.Text = "http://" + txtWebAdd.Text
                Else
                    txtWebAdd.Text = "http://www." + txtWebAdd.Text
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnHomeMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHomeMain.Click
        pnYaya.Visible = False
        pnYaya.SendToBack()
        WebBrowser1.Visible = True
        WebBrowser1.BringToFront()
        WebBrowser1.Navigate("http://www.google.com")
        txtWebAdd.Text = "http://www.google.com"
    End Sub

    Private Sub btnInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInfo.Click
        picRoom1.ImageLocation = Application.StartupPath & "\room4.jpg"
        If pnInfo = False Then
            WebBrowser1.Visible = False
            WebBrowser1.SendToBack()
            pnYaya.Visible = False
            pnYaya.SendToBack()
            pnInfo = True
        Else
            websiteDisplay()
            pnInfo = False
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pnYaya.Visible = False
        pnYaya.SendToBack()
        conDatabase = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Application.StartupPath & "\db.mdb")
        cmdDatabase = New OleDbCommand
        cmdDatabase.Connection = conDatabase
        Timer1.Enabled = True
        Timer1.Start()
    End Sub

    Private Sub txtWebAdd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWebAdd.KeyDown
        If e.KeyCode = Keys.Enter Then
            websiteDisplay()
        End If
    End Sub

    Private Sub btnFav_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFav.Click
        txtWebAdd.Text = "http://www.yayahotel.com"
        WebBrowser1.Visible = False
        WebBrowser1.SendToBack()
        pnYaya.Visible = True
        pnYaya.BringToFront()
    End Sub

    Private Sub btnRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegister.Click
        loginMode = "create"
        picBG.Visible = False
        picBG.SendToBack()
        pnProfile.Visible = True
        pnProfile.BringToFront()
        pnProfile.Location = New Point(397, 179)
        btnRegister.Enabled = False
        btnLogin.Enabled = False
        btnConfirmProfile.Text = "Register"
        txtLoginName.ReadOnly = False
    End Sub

    Private Sub btnConfirmProfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfirmProfile.Click
        If txtEmail.Text.Length < 1 Or txtLoginName.Text.Length < 1 Or txtPW1.Text.Length < 1 Or txtPW2.Text.Length < 1 Or txtFullName.Text.Length < 1 Or txtIC.Text.Length < 1 Or txtContact.Text.Length < 1 Or txtAdd.Text.Length < 1 Then
            MessageBox.Show("Please insert all details!", "Data Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            If Not txtPW1.Text = txtPW2.Text Then
                MessageBox.Show("Password NOT match!", "Password Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPW1.Text = Nothing
                txtPW2.Text = Nothing
                Exit Sub
            Else
                If Not txtEmail.Text = Nothing And txtEmail.Text.Contains("@") And txtEmail.Text.Substring(txtEmail.Text.Length - 3, 3) = "com" Then

                    cmdDatabase.CommandText = "INSERT INTO [user] ([username], [password],[fullname],[ic],[add],[contact],[email])" & "VALUES('" & txtLoginName.Text & "','" & txtPW1.Text & "','" & txtFullName.Text & "','" & txtIC.Text & "','" & txtAdd.Text & "','" & txtContact.Text & "','" & txtEmail.Text & "')"
                    conDatabase.Close()
                    If conDatabase.State = ConnectionState.Closed Then conDatabase.Open()
                    cmdDatabase.ExecuteNonQuery()
                    conDatabase.Close()
                    txtLoginName.Text = Nothing
                    txtPW1.Text = Nothing
                    txtPW2.Text = Nothing
                    txtFullName.Text = Nothing
                    txtIC.Text = Nothing
                    txtContact.Text = Nothing
                    txtAdd.Text = Nothing
                    btnLogin.Enabled = True
                    btnRegister.Enabled = True
                    pnProfile.Visible = False
                    pnProfile.SendToBack()
                    picBG.Visible = True
                    picBG.BringToFront()
                    MessageBox.Show("New user account created!", "Success!!!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Email NOT match!" + vbCr + "Please enter a valid email address.", "Email Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    Private Sub btnCancelProfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelProfile.Click
        txtLoginName.Text = Nothing
        txtPW1.Text = Nothing
        txtPW2.Text = Nothing
        txtFullName.Text = Nothing
        txtIC.Text = Nothing
        txtContact.Text = Nothing
        txtAdd.Text = Nothing
        btnLogin.Enabled = True
        btnRegister.Enabled = True
        pnProfile.Visible = False
        pnProfile.SendToBack()
        picBG.Visible = True
        picBG.BringToFront()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If btnLogin.Text = "Login" Then
            conDatabase.Close()
            cmdDatabase.CommandText = "SELECT * FROM [user]"
            If conDatabase.State = ConnectionState.Closed Then conDatabase.Open()
            Dim sdrICList As OleDbDataReader = cmdDatabase.ExecuteReader()
            While sdrICList.Read = True
                If txtUsername.Text = sdrICList.Item("username") Then
                    If txtPW.Text = sdrICList.Item("password") Then
                        loginSuccess = "yes"
                        email = sdrICList.Item("email")
                        Reservation.UserName = sdrICList.Item("username")
                    Else
                        loginSuccess = "pwfalse"
                        MessageBox.Show("Password Not Match!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtPW.Text = ""
                    End If
                End If
            End While
            sdrICList.Close()

            If loginSuccess = "yes" Then
                txtUsername.ReadOnly = True
                txtPW.ReadOnly = True
                btnLogin.Text = "Logout"
                btnRegister.Enabled = False
                picBG.Visible = False
                picBG.SendToBack()
                pnChooseRoom.Visible = True
                pnChooseRoom.BringToFront()
            ElseIf loginSuccess = "pwfalse" Then
            Else
                MessageBox.Show("Username Not Found!!!", "Error!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtUsername.Text = Nothing
                txtPW.Text = Nothing
            End If

            loginSuccess = Nothing
        Else
            txtUsername.Text = Nothing
            txtUsername.ReadOnly = False
            txtPW.Text = Nothing
            txtPW.ReadOnly = False
            btnLogin.Text = "Login"
            btnRegister.Enabled = True
            picBG.Visible = True
            picBG.BringToFront()
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
        End If
    End Sub

    Private Sub picBG_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles picBG.VisibleChanged
        picBG.Location = New Point(220, 91)
    End Sub

    Private Sub Panel4_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnChooseRoom.VisibleChanged
        pnChooseRoom.Location = New Point(220, 91)
    End Sub

    Private Sub room1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room1.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $150.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 150
            room = "Simple Double"
            confirmBooking()
        End If
    End Sub

    Private Sub room2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room2.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $200.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 200
            room = "Deluxe Double"
            confirmBooking()
        End If
    End Sub

    Private Sub room3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room3.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $250.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 250
            room = "Romance Double"
            confirmBooking()
        End If
    End Sub

    Private Sub room4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room4.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $300.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 300
            room = "Business Double"
            confirmBooking()
        End If
    End Sub

    Private Sub room5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room5.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $300.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 300
            room = "Business Double Classic"
            confirmBooking()
        End If
    End Sub

    Private Sub room6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room6.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $400.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 400
            room = "Luxury+ Double"
            confirmBooking()
        End If
    End Sub

    Private Sub room7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room7.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $550.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 550
            room = "Family Suite"
            confirmBooking()
        End If
    End Sub

    Private Sub room8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room8.Click
        Dim book As Integer

        book = MessageBox.Show("Room Price : $800.00", "Room Price", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If book = 1 Then
            pnChooseRoom.Visible = False
            pnChooseRoom.SendToBack()
            roomPrice = 800
            room = "President Suite"
            confirmBooking()
        End If
    End Sub

    Public Shared IsDateChoosen As Boolean
    Private Sub confirmBooking()
        Dim confirm As Integer
        Dim cardNumber As String = ""
        Dim bookDetails As String = ""

        tempEmail = email

        Dim chdTime As ChooseDateTime = New ChooseDateTime()
        chdTime.Show()
        chdTime.BringToFront()
        'Threading.Thread.Sleep(5000)

        'confirm = MessageBox.Show("You are going to reserve [" + room + "] with total price of $" + roomPrice.ToString, "Confirm?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        'If confirm = 6 Then
        '    cardNumber = InputBox("Please Key-in Your Credit Card Number to make payment", "Payment Gateway", "[CARD NUMBER HERE]")
        '    If IsNumeric(cardNumber) And cardNumber = "5167289400397530" Then
        '        System.Threading.Thread.Sleep(2000)
        '        MessageBox.Show("Payment Successful!", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '        bookDetails = "You room reservation for [" + room + "], is successful. Your reference Number is " + room.Substring(0, 1) + roomPrice.ToString + Now.Millisecond.ToString.Substring(0, 2) + "."
        '        email = InputBox("You booking reference number is " + room.Substring(0, 1) + roomPrice.ToString + Now.Millisecond.ToString.Substring(0, 2) + "." + vbCr _
        '                            + "Reservation details will be email to your email address." + vbCr _
        '                            + "Please Confirm Your Email Address as below : ", "EMAIL", email)

        '        If Not email = Nothing And email.Contains("@") Then
        '            If email.Substring(email.Length - 3, 3) = "com" Then
        '                Try

        '                    Mail.Subject = "[SUCCESS] YAYA Hotel Room Reservation."
        '                    Mail.From = New MailAddress("yahyashema@gmail.com")
        '                    SMTP.Credentials = New System.Net.NetworkCredential("yahyashema@gmail.com", "yayus12345")
        '                    Mail.To.Add("anuragc64@gmail.com")
        '                    Mail.Body = "Dear Value Customer," + vbCr + vbCr + bookDetails + vbCr + vbCr + "Thank you for choosing YAYA Hotel." + vbCr + vbCr + vbCr + "Warmly Regards," + _
        '                                vbCr + "YAYA Hotel Management Team"
        '                    SMTP.EnableSsl = True
        '                    SMTP.Port = "587"
        '                    SMTP.Send(Mail)

        '                    MessageBox.Show("Email Sent!!!", "Email Sent!!!", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '                    email = tempEmail
        '                    room = Nothing
        '                    roomPrice = 0
        '                    pnChooseRoom.Visible = True
        '                    pnChooseRoom.BringToFront()
        '                Catch ex As Exception
        '                    MessageBox.Show("Internet Connection Error! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '                    email = tempEmail
        '                    pnChooseRoom.Visible = True
        '                    pnChooseRoom.BringToFront()
        '                End Try
        '            Else
        '                MessageBox.Show("Email Address Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '                pnChooseRoom.Visible = True
        '                pnChooseRoom.BringToFront()
        '                email = tempEmail
        '            End If
        '        Else
        '            MessageBox.Show("Email Address Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '            pnChooseRoom.Visible = True
        '            pnChooseRoom.BringToFront()
        '            email = tempEmail
        '        End If
        '    Else
        '        MessageBox.Show("Card Number Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        pnChooseRoom.Visible = True
        '        pnChooseRoom.BringToFront()
        '        email = tempEmail
        '    End If
        'End If
        'email = tempEmail
    End Sub

    Public Sub confirmBooking1()
        Dim confirm As Integer
        Dim cardNumber As String = ""
        Dim bookDetails As String = ""

        tempEmail = email

        If (IsDateChoosen) Then
            confirm = MessageBox.Show("You are going to reserve [" + room + "] with total price of $" + roomPrice.ToString, "Confirm?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confirm = 6 Then
                cardNumber = InputBox("Please Key-in Your Credit Card Number to make payment", "Payment Gateway", "[CARD NUMBER HERE]")
                If IsNumeric(cardNumber) And cardNumber = "5167289400397530" Then
                    System.Threading.Thread.Sleep(2000)

                    bookDetails = "You room reservation for [" + room + "], is successful. Your reference Number is " + room.Substring(0, 1) + roomPrice.ToString + Now.Millisecond.ToString.Substring(0, 2) + "."
                    email = InputBox("You booking reference number is " + room.Substring(0, 1) + roomPrice.ToString + Now.Millisecond.ToString.Substring(0, 2) + "." + vbCr _
                                        + "Reservation details will be email to your email address." + vbCr _
                                        + "Please Confirm Your Email Address as below : ", "EMAIL", email)

                    Reservation.Room = room
                    Reservation.RoomPrice = roomPrice
                    Reservation.BookingTime = DateAndTime.Now.ToString

                    cmdDatabase.CommandText = "INSERT INTO [Booking] ([Username],[Checkin], [Checkout],[BookingTime],[Room],[RoomPrice])" & "VALUES('" & Reservation.UserName & "','" & Reservation.Checkin & "','" & Reservation.Checkout & "','" & Reservation.BookingTime & "','" & Reservation.Room & "','" & Reservation.RoomPrice & "')"
                    conDatabase.Close()
                    If conDatabase.State = ConnectionState.Closed Then conDatabase.Open()
                    cmdDatabase.ExecuteNonQuery()
                    conDatabase.Close()


                    If Not email = Nothing And email.Contains("@") Then
                        If email.Substring(email.Length - 3, 3) = "com" Then
                            Try

                                Mail.Subject = "[SUCCESS] YAYA Hotel Room Reservation."
                                Mail.From = New MailAddress("yahyashema@gmail.com")
                                SMTP.Credentials = New System.Net.NetworkCredential("yahyashema@gmail.com", "yayus12345")
                                Mail.To.Add(email)
                                Mail.Body = "Dear Value Customer," + vbCr + vbCr + bookDetails + vbCr + vbCr + "Thank you for choosing YAYA Hotel." + vbCr + vbCr + vbCr + "Warmly Regards," + _
                                            vbCr + "YAYA Hotel Management Team"
                                SMTP.EnableSsl = True
                                SMTP.Port = "587"
                                SMTP.Send(Mail)
                                MessageBox.Show("Payment Successful and Email Sent!!!", "Email Sent!!!", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                email = tempEmail
                                room = Nothing
                                roomPrice = 0
                                pnChooseRoom.Visible = True
                                pnChooseRoom.BringToFront()
                            Catch ex As Exception
                                MessageBox.Show("Internet Connection Error! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                email = tempEmail
                                pnChooseRoom.Visible = True
                                pnChooseRoom.BringToFront()
                            End Try
                        Else
                            MessageBox.Show("Email Address Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            pnChooseRoom.Visible = True
                            pnChooseRoom.BringToFront()
                            email = tempEmail
                        End If
                    Else
                        MessageBox.Show("Email Address Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        pnChooseRoom.Visible = True
                        pnChooseRoom.BringToFront()
                        email = tempEmail
                    End If
                Else
                    MessageBox.Show("Card Number Invalid! Booking and payment REVERSED!", "INVALID", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    pnChooseRoom.Visible = True
                    pnChooseRoom.BringToFront()
                    email = tempEmail
                End If
            End If
            email = tempEmail
        End If
        'Threading.Thread.Sleep(5000)


    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If k < 9 Then
            picRoom1.ImageLocation = Application.StartupPath & "\room" & k.ToString & ".jpg"
            k += 1
            If k = 9 Then
                k = 1
            End If
        End If
        If l > 0 Then
            picRoom2.ImageLocation = Application.StartupPath & "\room" & l.ToString & ".jpg"
            l -= 1
            If l = 0 Then
                l = 8
            End If
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim adminDash As AdminLogin = New AdminLogin()
        adminDash.Show()
    End Sub
End Class
