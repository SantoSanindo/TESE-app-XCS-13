Imports System.Data.SqlClient

Public Class frmMain
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Dir(My.Application.Info.DirectoryPath & "\Config.INI") = "" Then
            MsgBox("Config.INI is missing")
            End
        End If

        If Dir(My.Application.Info.DirectoryPath & "\Setup.INI") = "" Then
            MsgBox("Setup.INI is missing")
            End
        End If

        'ReadINI((My.Application.Info.DirectoryPath & "\Config.INI"))

        'Modbus
        frmMsg.Show()
        frmMsg.Text1.Text = "Establishing link wtih PLC..."
        Modbus.Show()
        Modbus.Hide()
        If Modbus.lbl_status.Text <> "Connected" Then
            Ethernet.BackColor = Color.Red
            Exit Sub
        End If
        frmMsg.Text1.Text = "Connection to PLC established"
        frmMsg.Hide()
        Ethernet.BackColor = Color.Green


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        frmDebug.Show()
        Me.Hide()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Modbus.ShowDialog()
    End Sub

    Private Sub Command3_Click(sender As Object, e As EventArgs) Handles Command3.Click
        If Command3.Text = "Print" Then
            Command3.Text = "No Print"
        Else
            Command3.Text = "Print"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmDatabase.ShowDialog()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Barcode_Comm.PortOpen = False
        frmSelect.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Modbus.tulisModbus(40107, 0)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = 1 Then
            If Not Modbus.tulisModbus(40092, 1) Then
                MessageBox.Show("Unable to communicate with PLC - %MW92")
                Exit Sub
            End If
        Else
            If Not Modbus.tulisModbus(40092, 0) Then
                MessageBox.Show("Unable to communicate with PLC - %MW92")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button4.Text = "Start" Then
            If Not Modbus.tulisModbus(40320, 1) Then
                MessageBox.Show("Unable to communicate with PLC - %MW320")
                Exit Sub
            End If
            Button4.Text = "Stop"
        ElseIf Button4.Text = "Stop" Then
            If Not Modbus.tulisModbus(40320, 0) Then
                MessageBox.Show("Unable to communicate with PLC - %MW320")
                Exit Sub
            End If
            Button4.Text = "Start"
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Button6.Text = "Door Unlock" Then
            If Not Modbus.tulisModbus(40260, 1) Then
                MessageBox.Show("Unable to communicate with PLC - %MW260")
                Exit Sub
            End If
            Button6.Text = "Door Lock"
        ElseIf Button6.Text = "Door Lock" Then
            If Not Modbus.tulisModbus(40260, 0) Then
                MessageBox.Show("Unable to communicate with PLC - %MW260")
                Exit Sub
            End If
            Button6.Text = "Door Unlock"
        End If
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        If Button10.Visible = False Then
            Button10.Visible = True
            Button8.Visible = True
        Else
            Button10.Visible = False
            Button8.Visible = False
        End If
    End Sub
    '######################################################################################################################
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'sub program belum dibuat
    End Sub
    '######################################################################################################################

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        frmMsg.Show()
        frmMsg.Text1.Text = "Homing Indexer..." & vbCrLf & "Please wait..."
        Modbus.tulisModbus(40401, 1)
        Do
            If Modbus.bacaModbus(40401) = 0 Then
                Exit Do
            End If
            If Modbus.bacaModbus(40190) <> 0 Then
                Exit Do
            End If
            System.Windows.Forms.Application.DoEvents()
        Loop
        frmMsg.Close()
        Timer2.Enabled = False
        Timer1.Enabled = True
    End Sub
End Class
