Public Class frmDebug
    Private Sub cmd_back_Click(sender As Object, e As EventArgs) Handles cmd_back.Click
        frmMain.Show()
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox1.Text = CStr(Modbus.bacaModbus(40102))
        TextBox2.Text = CStr(Modbus.bacaModbus(40101))
        TextBox3.Text = CStr(Modbus.bacaModbus(40122))
        TextBox4.Text = CStr(Modbus.bacaModbus(40123))
        TextBox5.Text = CStr(Modbus.bacaModbus(40125))
        TextBox6.Text = CStr(Modbus.bacaModbus(40129))
        TextBox7.Text = CStr(Modbus.bacaModbus(40124))
        TextBox9.Text = CStr(Modbus.bacaModbus(40142))
        TextBox10.Text = CStr(Modbus.bacaModbus(40143))
        TextBox11.Text = CStr(Modbus.bacaModbus(40145))
        TextBox12.Text = CStr(Modbus.bacaModbus(40149))
        TextBox13.Text = CStr(Modbus.bacaModbus(40144))
    End Sub

    Private Sub TextBox16_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox16.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40110, CInt(TextBox16.Text))
        End If
    End Sub

    Private Sub TextBox18_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox18.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40130, CInt(TextBox18.Text))
        End If
    End Sub

    Private Sub TextBox20_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox20.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40133, CInt(TextBox20.Text))
        End If
    End Sub

    Private Sub TextBox22_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox22.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40131, CInt(TextBox22.Text))
        End If
    End Sub

    Private Sub TextBox24_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox24.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40134, CInt(TextBox24.Text))
        End If
    End Sub

    Private Sub TextBox27_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox27.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40150, CInt(TextBox27.Text))
        End If
    End Sub

    Private Sub TextBox28_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox28.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40153, CInt(TextBox28.Text))
        End If
    End Sub

    Private Sub TextBox29_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox29.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40151, CInt(TextBox29.Text))
        End If
    End Sub

    Private Sub TextBox30_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox30.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40154, CInt(TextBox30.Text))
        End If
    End Sub

    Private Sub TextBox34_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox34.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40197, CInt(TextBox34.Text))
        End If
    End Sub

    Private Sub TextBox35_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox35.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40199, CInt(TextBox35.Text))
        End If
    End Sub

    Private Sub TextBox36_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox36.KeyPress
        If e.KeyChar = Chr(13) Then
            Modbus.tulisModbus(40198, CInt(TextBox36.Text))
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Modbus.tulisModbus(CInt(TextBox38.Text), CInt(TextBox39.Text))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmMain.Chroma_Comm.Write(TextBox37.Text & vbCrLf)
    End Sub
End Class