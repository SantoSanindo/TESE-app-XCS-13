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
    End Sub
End Class