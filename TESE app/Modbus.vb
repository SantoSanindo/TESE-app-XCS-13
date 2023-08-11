Imports EasyModbus

Public Class Modbus
    Dim modbusClient As ModbusClient

    Public Function bacaModbus(addrs As Integer) As Integer
        If lbl_status.Text = "Connected" Then
            txtAddress.Text = addrs
            read_modbus()
            Return CInt(txtValue.Text)
        Else
            MsgBox("modbus Not connected")
            Return 0
        End If
    End Function

    Public Function tulisModbus(addrs As Integer, value As Integer) As Boolean
        If lbl_status.Text = "Connected" Then
            txtAddress.Text = addrs
            txtNewValue.Text = value
            write_modbus()
            Return True
        Else
            MsgBox("modbus Not connected")
            Return False
        End If
    End Function

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            modbusClient = New ModbusClient(txtIP.Text, Val(txtPort.Text))
            modbusClient.Connect()

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            lbl_status.Text = "Connected"
            lbl_status.ForeColor = Color.Green
        Catch ex As Exception
            MsgBox("Modbus Error Connect! " & ex.Message)
        End Try

    End Sub

    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        Try
            modbusClient.Disconnect()

            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Modbus Error Disconnect! " & ex.Message)
        End Try
        lbl_status.Text = "Not Connected"
        lbl_status.ForeColor = Color.Red
    End Sub

    Private Sub read_modbus()
        Dim StartAddress As Integer = Val(txtAddress.Text)

        Select Case cboRegType.SelectedIndex
            Case 0
                Dim ReadValues() As Boolean = modbusClient.ReadCoils(StartAddress, 1)
                txtValue.Text = ReadValues(0)
            Case 1
                If StartAddress > 10000 Then StartAddress = StartAddress - 10001
                Dim ReadValues() As Boolean = modbusClient.ReadDiscreteInputs(StartAddress, 1)
                txtValue.Text = ReadValues(0)
            Case 2
                If StartAddress > 30000 Then StartAddress = StartAddress - 30001
                Dim ReadValues() As Integer = modbusClient.ReadInputRegisters(StartAddress, 1)
                txtValue.Text = ReadValues(0)
            Case 3
                If StartAddress > 40000 Then StartAddress = StartAddress - 40001
                Dim ReadValues() As Integer = modbusClient.ReadHoldingRegisters(StartAddress, 1)
                txtValue.Text = ReadValues(0)
        End Select
    End Sub

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click
        Try
            read_modbus()

        Catch ex As Exception
            MsgBox("Modbus Error Read! " & ex.Message)
        End Try
    End Sub

    Private Sub write_modbus()
        Dim StartAddress As Integer = Val(txtAddress.Text)

        Select Case cboRegType.SelectedIndex
            Case 0
                Dim WriteVals(0) As Boolean
                WriteVals(0) = Val(txtNewValue.Text)
                modbusClient.WriteMultipleCoils(StartAddress, WriteVals)
            Case 1
                    '
            Case 2
                    '
            Case 3
                If StartAddress > 40000 Then StartAddress = StartAddress - 40001
                Dim WriteVals(0) As Integer
                WriteVals(0) = Val(txtNewValue.Text)
                modbusClient.WriteMultipleRegisters(StartAddress, WriteVals)
        End Select
    End Sub

    Private Sub btnWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWrite.Click
        Try
            write_modbus()

        Catch ex As Exception
            MsgBox("Modbus Error Write! " & ex.Message)
        End Try
    End Sub

    Private Sub cboRegType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRegType.SelectedIndexChanged
        Dim iSel As Integer = cboRegType.SelectedIndex
        If iSel > -1 Then
            If iSel = 0 Then
                txtAddress.Text = "00001"
                btnWrite.Enabled = True
            End If
            If iSel = 1 Then
                txtAddress.Text = "10001"
                btnWrite.Enabled = False
            End If

            If iSel = 2 Then
                txtAddress.Text = "30001"
                btnWrite.Enabled = False
            End If
            If iSel = 3 Then
                txtAddress.Text = "40001"
                btnWrite.Enabled = True
            End If

        End If
    End Sub

    Private Sub Modbus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboRegType.SelectedIndex = 3
        btnConnect.PerformClick()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
End Class