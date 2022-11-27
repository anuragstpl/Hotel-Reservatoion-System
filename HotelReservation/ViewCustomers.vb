Imports System.Data.OleDb

Public Class ViewCustomers
    Dim conDatabase As OleDbConnection
    Dim cmdDatabase As OleDbCommand = New OleDbCommand

    Private Sub ViewCustomers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conDatabase = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" & Application.StartupPath & "\db.mdb")
        cmdDatabase = New OleDbCommand
        cmdDatabase.Connection = conDatabase
        cmdDatabase.CommandText = "Select * from [user];"
        If conDatabase.State = ConnectionState.Closed Then conDatabase.Open()
        Dim ds As DataSet = New DataSet()
        Dim old As OleDbDataAdapter = New OleDbDataAdapter(cmdDatabase)
        old.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        'DataGridView1.DataBind()
        cmdDatabase.ExecuteNonQuery()
        conDatabase.Close()
    End Sub
End Class