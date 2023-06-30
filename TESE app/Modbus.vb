Imports EasyModbus

Public Class Modbus
    Dim modbusClient As ModbusClient
    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            modbusClient = New ModbusClient(txtIP.Text, Val(txtPort.Text))
            modbusClient.Connect()

            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
        Catch ex As Exception
            MsgBox("Error Connect! " & ex.Message)
        End Try

    End Sub

    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        Try
            modbusClient.Disconnect()

            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error Disconnect! " & ex.Message)
        End Try
    End Sub

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click
        Try
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

        Catch ex As Exception
            MsgBox("Error Read! " & ex.Message)
        End Try
    End Sub

    Private Sub btnWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWrite.Click
        Try
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

        Catch ex As Exception
            MsgBox("Error Write! " & ex.Message)
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
End Class