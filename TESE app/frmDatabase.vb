Imports System.Data.SqlClient
Public Class frmDatabase
    Private Sub frmDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 1 To 16
            ComboBox3.Items.Add(i)
        Next


        ComboBox1.Items.Add("Plastic")
        ComboBox1.Items.Add("Zamak")

        ComboBox5.Items.Add("Energisation")
        ComboBox5.Items.Add("De-Energisation")

        ComboBox3.Items.Add("24VDC")
        ComboBox3.Items.Add("24VAC")
        ComboBox3.Items.Add("120VAC")
        ComboBox3.Items.Add("230VAC")

        Dim query = "SELECT ModelName FROM TESE.dbo.XCS_13_Parameter$"
        Dim dt = KoneksiDB.bacaData(query).Tables(0)

        For i As Integer = 0 To dt.Rows.Count - 1
            ComboBox2.Items.Add(dt.Rows(i).Item(0))
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        frmMain.Show()
    End Sub
    Private Sub ClearRefCombo()
        ComboBox2.Items.Clear()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "" Then Exit Sub
        Dim query = "SELECT * FROM TESE.dbo.XCS_13_Parameter$ WHERE ModelName = '" & ComboBox2.Text & "'"
        Dim dt = KoneksiDB.bacaData(query).Tables(0)

        Txt_Partnos.Text = dt.Rows(0).Item("ArticleNos")
        ComboBox1.Text = dt.Rows(0).Item("MaterialType")
        ComboBox3.Text = dt.Rows(0).Item("TensionType")
        ComboBox5.Text = dt.Rows(0).Item("FunctionType")
        ComboBox40.Text = dt.Rows(0).Item("Contact1Type")
        ComboBox41.Text = dt.Rows(0).Item("Contact2Type")
        ComboBox42.Text = dt.Rows(0).Item("Contact3Type")
        ComboBox43.Text = dt.Rows(0).Item("Contact4Type")
        ComboBox44.Text = dt.Rows(0).Item("Contact5Type")
        ComboBox45.Text = dt.Rows(0).Item("Contact6Type")
        ComboBox60.Text = dt.Rows(0).Item("Contact1_W_Key")
        ComboBox61.Text = dt.Rows(0).Item("Contact2_W_Key")
        ComboBox62.Text = dt.Rows(0).Item("Contact3_W_Key")
        ComboBox63.Text = dt.Rows(0).Item("Contact4_W_Key")
        ComboBox64.Text = dt.Rows(0).Item("Contact5_W_Key")
        ComboBox65.Text = dt.Rows(0).Item("Contact6_W_Key")
        ComboBox70.Text = dt.Rows(0).Item("Contact1_W_Key_Ten")
        ComboBox71.Text = dt.Rows(0).Item("Contact2_W_Key_Ten")
        ComboBox72.Text = dt.Rows(0).Item("Contact3_W_Key_Ten")
        ComboBox73.Text = dt.Rows(0).Item("Contact4_W_Key_Ten")
        ComboBox74.Text = dt.Rows(0).Item("Contact5_W_Key_Ten")
        ComboBox75.Text = dt.Rows(0).Item("Contact6_W_Key_Ten")
    End Sub

    Private Sub cmdsave_Click(sender As Object, e As EventArgs) Handles cmdsave.Click
        Dim konek As New SqlConnection(database)

        Dim sql As String = "UPDATE TESE.dbo.XCS_13_Parameter$ SET ArticleNos = '" & Txt_Partnos.Text & "', MaterialType= '" & ComboBox1.Text & "',TensionType = '" & ComboBox3.Text & "', FunctionType= '" & ComboBox5.Text & "', Contact1Type = '" & ComboBox40.Text & "', Contact2Type= '" & ComboBox41.Text & "', Contact3Type = '" & ComboBox42.Text & "', Contact4Type= '" & ComboBox43.Text & "', Contact5Type = '" & ComboBox44.Text & "', Contact6Type= '" & ComboBox45.Text & "', Contact1_W_Key = '" & ComboBox60.Text & "',Contact2_W_Key= '" & ComboBox61.Text & "', Contact3_W_Key = '" & ComboBox62.Text & "',Contact4_W_Key= '" & ComboBox63.Text & "', Contact5_W_Key= '" & ComboBox64.Text & "',Contact6_W_Key = '" & ComboBox65.Text & "', Contact1_W_Key_Ten= '" & ComboBox70.Text & "',Contact2_W_Key_Ten = '" & ComboBox71.Text & "', Contact3_W_Key_Ten = '" & ComboBox72.Text & "',Contact4_W_Key_Ten = '" & ComboBox73.Text & "', Contact5_W_Key_Ten = '" & ComboBox74.Text & "',Contact6_W_Key_Ten = '" & ComboBox75.Text & "' Where ModelName = '" & ComboBox2.Text & "'"

        konek.Open()
        Dim sc As New SqlCommand(sql, konek)
        Dim adapter As New SqlDataAdapter(sc)
        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            MsgBox("Saving succeed!")
            konek.Close()
        Else
            MsgBox("Saving Failed!")
        End If

    End Sub
End Class