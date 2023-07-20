Public Class frmInterrupt
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "Continue" Then
            Modbus.tulisModbus(40190, 49)
            Me.Close()
        ElseIf Button1.Text = "Reset" Then
            Modbus.tulisModbus(40190, 5)
            Me.Close()
        End If
    End Sub
End Class