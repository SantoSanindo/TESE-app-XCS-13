Imports System.Data.SqlClient
Imports TESE_app.KoneksiDB

Public Class testing
    Private Sub testing_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim query = "SELECT * FROM tese.dbo.[Login]"
        DataGridView1.DataSource = KoneksiDB.bacaData(query).Tables(0)
    End Sub



End Class
