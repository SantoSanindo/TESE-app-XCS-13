Imports System.Data.SqlClient

Public Class frmMain
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Modbus.Show()
    End Sub

    Private Sub Command3_Click(sender As Object, e As EventArgs) Handles Command3.Click
        If Command3.Text = "Print" Then
            Command3.Text = "No Print"
        Else
            Command3.Text = "Print"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        frmDatabase.Show()
    End Sub
End Class
