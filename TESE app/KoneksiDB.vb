
Imports System.Data.SqlClient
Module KoneksiDB
    Public Function bacaData(query As String) As DataSet

        Dim database As String = "Data Source=192.168.1.11\SQLEXPRESS;
            initial catalog=TESE;
            Persist Security Info=True;
            User ID=tese;
            Password=Sanindo123;
            Connect Timeout=15000;
            Max Pool Size=15000;
            Pooling=True"
        Dim koneksi As New SqlConnection(database)


        query = "SELECT * FROM tese.dbo.[Login]"

        Dim sc As New SqlCommand(query, koneksi)
        Dim adapter As New SqlDataAdapter(sc)
        Dim ds As New DataSet

        adapter.Fill(ds)
        koneksi.Close()
        Return ds

    End Function


End Module
