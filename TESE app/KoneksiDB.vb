
Imports System.Data.SqlClient
Module KoneksiDB

    Public Function koneksi() As SqlConnection
        Dim database As String = "Data Source=192.168.1.11\SQLEXPRESS;
            initial catalog=TESE;
            Persist Security Info=True;
            User ID=tese;
            Password=Sanindo123;
            Connect Timeout=15000;
            Max Pool Size=15000;
            Pooling=True"
        Dim konek As New SqlConnection(database)
        Return konek
    End Function

    Public Function bacaData(query As String) As DataSet
        Try

            Dim sc As New SqlCommand(query, koneksi)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            'koneksi.Close()
            Return ds
        Catch ex As Exception
            MsgBox("Database connection Error!")
        End Try

    End Function


End Module
