Imports System.Data.SqlClient
Imports System.Threading

Public Class frmMain
    Dim AssyBuf As String
    Dim Smacbuf As String
    Dim ChromaBuf As String
    Dim ChromaFBbuf As String
    Dim TempPSN As String
    Public S1_cavitynos As Integer
    Dim SychResult As String
    Dim EnableCount As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button4.Text = "Stop" Then
            MsgBox("Please set to Stop Mode before quiting the program")
            Exit Sub
        Else
            CloseCodesoft()
            Barcode_Comm.Close()
            RFID_Comm.Close()
            SMAC_Comm.Close()
            Chroma_Comm.Close()
            End
        End If
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

        ReadINI((My.Application.Info.DirectoryPath & "\Config.INI"))

        'Modbus
        frmMsg.Show()
        frmMsg.Text1.Text = "Establishing link with PLC..."
        Modbus.Show()
        Modbus.Hide()
        If Modbus.lbl_status.Text <> "Connected" Then
            Ethernet.BackColor = Color.Red
            End
        End If
        frmMsg.Text1.Text = "Connection to PLC established"
        'frmMsg.Hide()
        Ethernet.BackColor = Color.Green

        Chroma_Comm.Open()
        frmMsg.Text1.Text = "Initalizing Chroma 19053..."
        If Not Init_Chroma() Then
            frmMsg.Text1.Text = "Unable to initalize chroma" & vbCrLf
            End
        End If
        GetLastConfig()
        frmMsg.Text1.Text = "Loading Parameter..."
        If Not LoadParameter(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load database for last working reference")
            End
        End If
        lbl_WOnos.Text = LoadWOfrRFID.JobNos
        lbl_currentref.Text = LoadWOfrRFID.JobModelName
        lbl_wocounter.Text = CStr(LoadWOfrRFID.JobQTy)
        lbl_tagnos.Text = LoadWOfrRFID.JobRFIDTag
        Label8.Text = CStr(LoadWOfrRFID.JobUnitaryCount)
        lbl_ArticleNos.Text = LoadWOfrRFID.JobArticle
        frmMsg.Text1.Text = "DownLoading Parameter to PLC..."
        If Not DownloadParameter() Then
            MsgBox("Unable to communicate with PLC")
            End
        End If

        frmMsg.Text1.Text = "Loading CodeSoft..."
        OpenCodesoft()
        frmMsg.Text1.Text = "Loading TLP2844Z..."
        If Not SetPrinter("Zebra TLP2844-Z", "USB001") Then
            frmMsg.Text1.Text = "Unable to switch Printer" & vbCrLf
            End
        End If

        frmMsg.Text1.Text = "Verifying Incoming air supply..."
        'If Modbus.bacaModbus(40090) = 0 Then
        '    frmMsg.Text1.Text = "No Incoming air supply"
        '    End
        'End If


        SMAC_Comm.Open()
        RFID_Comm.Open()
        Barcode_Comm.Open()

        SMAC_Comm.Write(Chr(27))
        Thread.Sleep(10)
        SMAC_Comm.Write("ms0" & Chr(13))

        If Modbus.bacaModbus(40320) = 1 Then
            Button4.Text = "Stop"
        Else
            Button4.Text = "Start"
        End If
        Timer2.Enabled = True
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
        Barcode_Comm.Close()
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        GetLastConfig()
        frmMsg.Show()
        frmMsg.Text1.Text = "Loading Parameter..."
        If Not LoadParameter(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load database for last working reference")
            End
        End If
        lbl_WOnos.Text = LoadWOfrRFID.JobNos
        lbl_currentref.Text = LoadWOfrRFID.JobModelName
        lbl_wocounter.Text = CStr(LoadWOfrRFID.JobQTy)
        lbl_tagnos.Text = LoadWOfrRFID.JobRFIDTag
        Label8.Text = CStr(LoadWOfrRFID.JobUnitaryCount)
        frmMsg.Text1.Text = "DownLoading Parameter to PLC..."
        If Not DownloadParameter() Then
            frmMsg.Text1.Text = "Unable to communicate with PLC"
            Thread.Sleep(10)
            End
        End If
        frmMsg.Close()
    End Sub

    Private Function LoadParameter(UnitRef As String) As Boolean
        Try
            If UnitRef = "" Then
                Return False
                Exit Function
            End If

            Dim query = "SELECT * FROM TESE.dbo.XCS_13_Parameter$ WHERE ModelName = '" & UnitRef & "'"
            Dim dt = KoneksiDB.bacaData(query).Tables(0)

            Parameter.UnitTension = dt.Rows(0).Item("TensionType")
            Parameter.UnitPartNos = dt.Rows(0).Item("ArticleNos")
            LoadWOfrRFID.JobArticle = dt.Rows(0).Item("ArticleNos")
            Parameter.UnitType = dt.Rows(0).Item("MaterialType")
            Parameter.UnitFunction = dt.Rows(0).Item("FunctionType")
            Parameter.UnitSycLL = dt.Rows(0).Item("SycLL")
            Parameter.UnitSycUL = dt.Rows(0).Item("SycUL")
            Parameter.UnitContact1_WO_Trig = dt.Rows(0).Item("Contact1Type")
            If Parameter.UnitContact1_WO_Trig = "None" Then
                Parameter.UnitContact1_WO_Trig = "0"
            End If

            Parameter.UnitContact2_WO_Trig = dt.Rows(0).Item("Contact2Type")
            If Parameter.UnitContact2_WO_Trig = "None" Then
                Parameter.UnitContact2_WO_Trig = "0"
            End If

            Parameter.UnitContact3_WO_Trig = dt.Rows(0).Item("Contact3Type")
            If Parameter.UnitContact3_WO_Trig = "None" Then
                Parameter.UnitContact3_WO_Trig = "0"
            End If

            Parameter.UnitContact4_WO_Trig = dt.Rows(0).Item("Contact4Type")
            If Parameter.UnitContact4_WO_Trig = "None" Then
                Parameter.UnitContact4_WO_Trig = "0"
            End If

            Parameter.UnitContact5_WO_Trig = dt.Rows(0).Item("Contact5Type")
            If Parameter.UnitContact5_WO_Trig = "None" Then
                Parameter.UnitContact5_WO_Trig = "0"
            End If

            Parameter.UnitContact6_WO_Trig = dt.Rows(0).Item("Contact6Type")
            If Parameter.UnitContact6_WO_Trig = "None" Then
                Parameter.UnitContact6_WO_Trig = "0"
            End If

            Parameter.UnitContact1_W_Key = dt.Rows(0).Item("Contact1_W_Key")
            If Parameter.UnitContact1_W_Key = "None" Then
                Parameter.UnitContact1_W_Key = "0"
            End If

            Parameter.UnitContact2_W_Key = dt.Rows(0).Item("Contact2_W_Key")
            If Parameter.UnitContact2_W_Key = "None" Then
                Parameter.UnitContact2_W_Key = "0"
            End If

            Parameter.UnitContact3_W_Key = dt.Rows(0).Item("Contact3_W_Key")
            If Parameter.UnitContact3_W_Key = "None" Then
                Parameter.UnitContact3_W_Key = "0"
            End If

            Parameter.UnitContact4_W_Key = dt.Rows(0).Item("Contact4_W_Key")
            If Parameter.UnitContact4_W_Key = "None" Then
                Parameter.UnitContact4_W_Key = "0"
            End If

            Parameter.UnitContact5_W_Key = dt.Rows(0).Item("Contact5_W_Key")
            If Parameter.UnitContact5_W_Key = "None" Then
                Parameter.UnitContact5_W_Key = "0"
            Else
                Parameter.UnitContact5_W_Key = dt.Rows(0).Item("Contact5_W_Key")
            End If

            Parameter.UnitContact6_W_Key = dt.Rows(0).Item("Contact6_W_Key")
            If Parameter.UnitContact6_W_Key = "None" Then
                Parameter.UnitContact6_W_Key = "0"
            End If
            Parameter.UnitContact1_W_Key_Ten = dt.Rows(0).Item("Contact1_W_Key_Ten")
            If Parameter.UnitContact1_W_Key_Ten = "None" Then
                Parameter.UnitContact1_W_Key_Ten = "0"
            End If
            Parameter.UnitContact2_W_Key_Ten = dt.Rows(0).Item("Contact2_W_Key_Ten")
            If Parameter.UnitContact2_W_Key_Ten = "None" Then
                Parameter.UnitContact2_W_Key_Ten = "0"
            End If
            Parameter.UnitContact3_W_Key_Ten = dt.Rows(0).Item("Contact3_W_Key_Ten")
            If Parameter.UnitContact3_W_Key_Ten = "None" Then
                Parameter.UnitContact3_W_Key_Ten = "0"
            End If
            Parameter.UnitContact4_W_Key_Ten = dt.Rows(0).Item("Contact4_W_Key_Ten")
            If Parameter.UnitContact4_W_Key_Ten = "None" Then
                Parameter.UnitContact4_W_Key_Ten = "0"
            End If
            Parameter.UnitContact5_W_Key_Ten = dt.Rows(0).Item("Contact5_W_Key_Ten")
            If Parameter.UnitContact5_W_Key_Ten = "None" Then
                Parameter.UnitContact5_W_Key_Ten = "0"
            End If
            Parameter.UnitContact6_W_Key_Ten = dt.Rows(0).Item("Contact6_W_Key_Ten")
            If Parameter.UnitContact6_W_Key_Ten = "None" Then
                Parameter.UnitContact6_W_Key_Ten = "0"
            End If

            Dim query2 = "SELECT * FROM TESE.dbo.XCS_13_Label$ WHERE ModelName = '" & UnitRef & "'"
            Dim dt2 = KoneksiDB.bacaData(query2).Tables(0)

            Parameter.UnitLabelTemplate = dt.Rows(0).Item("Schematic_Template")
            Parameter.UnitLabelPhoto = dt.Rows(0).Item("Schematic_Img")

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function DownloadParameter() As Boolean
        If Parameter.UnitFunction = "Energisation" Then
            If Not Modbus.tulisModbus(40209, 1) Then
                MsgBox("Unable to write to PLC - %MW209")
                Return False
                Exit Function
            End If
        Else
            If Not Modbus.tulisModbus(40209, 2) Then
                MsgBox("Unable to write to PLC - %MW209")
                Return False
                Exit Function
            End If
        End If

        If Parameter.UnitTension = "24VDC" Then
            If Not Modbus.tulisModbus(40210, 8) Then
                MsgBox("Unable to write to PLC - %MW210")
                Return False
                Exit Function
            End If
        ElseIf Parameter.UnitTension = "24VAC" Then
            If Not Modbus.tulisModbus(40210, 1) Then
                MsgBox("Unable to write to PLC - %MW210")
                Return False
                Exit Function
            End If
        ElseIf Parameter.UnitTension = "120VAC" Then
            If Not Modbus.tulisModbus(40210, 2) Then
                MsgBox("Unable to write to PLC - %MW210")
                Return False
                Exit Function
            End If
        ElseIf Parameter.UnitTension = "230VAC" Then
            If Not Modbus.tulisModbus(40210, 4) Then
                MsgBox("Unable to write to PLC - %MW210")
                Return False
                Exit Function
            End If
        End If

        If Parameter.UnitType = "Zamak" Then
            If Not Modbus.tulisModbus(40208, 1) Then
                MsgBox("Unable to write to PLC - %MW208")
                Return False
                Exit Function
            End If
        ElseIf Parameter.UnitType = "Plastic" Then
            If Not Modbus.tulisModbus(40208, 2) Then
                MsgBox("Unable to write to PLC - %MW208")
                Return False
                Exit Function
            End If
        End If

        Parameter.UnitContact_WO_Trig = CLng(BIN2Dec(Parameter.UnitContact6_WO_Trig & Parameter.UnitContact5_WO_Trig & Parameter.UnitContact4_WO_Trig & Parameter.UnitContact3_WO_Trig & Parameter.UnitContact2_WO_Trig & Parameter.UnitContact1_WO_Trig))
        Dim TempState As String
        TempState = Dec2Bin(Parameter.UnitContact_WO_Trig)
        Label100.Text = Mid(TempState, 16, 1)
        Label101.Text = Mid(TempState, 15, 1)
        Label102.Text = Mid(TempState, 14, 1)
        Label103.Text = Mid(TempState, 13, 1)
        Label104.Text = Mid(TempState, 12, 1)
        Label105.Text = Mid(TempState, 11, 1)
        If Not Modbus.tulisModbus(40201, Parameter.UnitContact_WO_Trig) Then
            MsgBox("Unable to write to PLC - %MW211")
            Return False
            Exit Function
        End If

        Parameter.UnitContact_W_Key = CLng(BIN2Dec(Parameter.UnitContact6_W_Key & Parameter.UnitContact5_W_Key & Parameter.UnitContact4_W_Key & Parameter.UnitContact3_W_Key & Parameter.UnitContact2_W_Key & Parameter.UnitContact1_W_Key))
        TempState = Dec2Bin(Parameter.UnitContact_W_Key)
        Label120.Text = Mid(TempState, 16, 1)
        Label121.Text = Mid(TempState, 15, 1)
        Label122.Text = Mid(TempState, 14, 1)
        Label123.Text = Mid(TempState, 13, 1)
        Label124.Text = Mid(TempState, 12, 1)
        Label125.Text = Mid(TempState, 11, 1)
        If Not Modbus.tulisModbus(40202, Parameter.UnitContact_W_Key) Then
            MsgBox("Unable to write to PLC - %MW212")
            Return False
            Exit Function
        End If

        If Label123.Text = "1" Then
            If Not Modbus.tulisModbus(40204, 1) Then
                MsgBox("Unable to write to PLC - %MW204")
                Return False
                Exit Function
            End If
        Else
            If Not Modbus.tulisModbus(40204, 0) Then
                MsgBox("Unable to write to PLC - %MW204")
                Return False
                Exit Function
            End If
        End If
        If Label124.Text = "1" Then
            If Not Modbus.tulisModbus(40205, 1) Then
                MsgBox("Unable to write to PLC - %MW205")
                Return False
                Exit Function
            End If
        Else
            If Not Modbus.tulisModbus(40205, 0) Then
                MsgBox("Unable to write to PLC - %MW205")
                Return False
                Exit Function
            End If
        End If
        If Label125.Text = "1" Then
            If Not Modbus.tulisModbus(40206, 1) Then
                MsgBox("Unable to write to PLC - %MW206")
                Return False
                Exit Function
            End If
        Else
            If Not Modbus.tulisModbus(40206, 0) Then
                MsgBox("Unable to write to PLC - %MW206")
                Return False
                Exit Function
            End If
        End If

        Parameter.UnitContact_W_Key_Ten = CInt(BIN2Dec(Parameter.UnitContact6_W_Key_Ten & Parameter.UnitContact5_W_Key_Ten & Parameter.UnitContact4_W_Key_Ten & Parameter.UnitContact3_W_Key_Ten & Parameter.UnitContact2_W_Key_Ten & Parameter.UnitContact1_W_Key_Ten))
        TempState = Dec2Bin(CDbl(Parameter.UnitContact_W_Key_Ten))
        Label130.Text = Mid(TempState, 16, 1)
        Label131.Text = Mid(TempState, 15, 1)
        Label132.Text = Mid(TempState, 14, 1)
        Label133.Text = Mid(TempState, 13, 1)
        Label134.Text = Mid(TempState, 12, 1)
        Label135.Text = Mid(TempState, 11, 1)
        If Not Modbus.tulisModbus(40203, Parameter.UnitContact_W_Key_Ten) Then
            MsgBox("Unable to write to PLC - %MW213")
            Return False
            Exit Function
        End If

        Return True
    End Function

    Public Sub PrintLabel(PSN As String)
        ActiveDoc.Variables.FormVariables.Item("Var14").Value = INILABELIMGPATH & Parameter.UnitLabelPhoto
        ActiveDoc.Variables.FormVariables.Item("Var1").Value = PSN & "X13"

        If PrintLab(1) = False Then
            MsgBox("Error. Can't to print...", MsgBoxStyle.Critical)
        End If
    End Sub

    Public Sub PrintFailure(PSN As String, ErrorCode As String)
        ActiveDoc.Variables.FormVariables.Item("Var13").Value = "Failure Code : " & ErrorCode
        ActiveDoc.Variables.FormVariables.Item("Var1").Value = PSN
        ActiveDoc.Variables.FormVariables.Item("Var10").Value = FailureMode(Val(ErrorCode))

        If PrintLab(1) = False Then
            MsgBox("Error. Can't to print...", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Function FailureMode(TestCode As Integer) As String
        Select Case TestCode
            Case 21
                Return "HiPot Selftest - Original State"
            Case 22
                Return "HiPot Selftest - State with key"
            Case 23
                Return "Hipot Selftest - State with key & tension"
            Case 31
                Return "Hipot - Test 1"
            Case 32
                Return "Hipot - Test 2"
            Case 33
                Return "Hipot - Test 3"
            Case 34
                Return "Hipot - Test 4"
            Case 35
                Return "Hipot - Test 5"
            Case 36
                Return "Hipot - Test 6"
            Case 37
                Return "Hipot - Test 7"
            Case 38
                Return "Hipot - Test 8"
            Case 71
                Return "Functional - Original State"
            Case 72
                Return "Functional - State with key"
            Case 73
                Return "Functional - Pull Test"
            Case 74
                Return "Functional - State with Key & tension"
            Case 75
                Return "Functional - LED Test"
            Case 76
                Return "Functional - Original State after test"
            Case 101
                Return "Sychronization - Original State"
            Case 102
                Return "Sychronization - Mechanism screw"
            Case 103
                Return "Sychronization - Measurement out of limits"
            Case 104
                Return "Sychronization - Timeout"
        End Select
    End Function

    '****************************************** Barcode_Comunication ******************************************
    Private Sub Barcode_Comm_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles Barcode_Comm.DataReceived
        AssyBuf = Barcode_Comm.ReadExisting()
        If InStr(1, AssyBuf, vbCrLf) <> 0 Then
            Me.Invoke(Sub()
                          Img_Result.Visible = False
                          Label79.Text = ""
                          Text3.Text = ""
                          lbl_msg.Text = ""
                          'Check if the Work Order Quantity is reached
                          If Val(Label8.Text) >= LoadWOfrRFID.JobQTy Then
                              Text3.Text = "WORK ORDER COMPLETED"
                              AssyBuf = ""
                              Exit Sub
                          End If

                          AssyBuf = Mid(AssyBuf, 1, InStr(1, AssyBuf, vbCr) - 1)
                          'Exit Sub

                          If Microsoft.VisualBasic.Left(AssyBuf, 6) <> LoadWOfrRFID.JobArticle Then
                              Text3.Text = "--> PSN - " & AssyBuf & " = wrong reference"
                              AssyBuf = ""
                              Exit Sub
                          Else
                              Text3.Text = "PSN - " & AssyBuf & vbCrLf
                          End If

                          Text3.Text = Text3.Text & "Loading " & AssyBuf & ".Txt..." & vbCrLf
                          If Dir(INIPSNFOLDERPATH & AssyBuf & ".Txt") = "" Then
                              Text3.Text = Text3.Text & "--> Unable to find" & AssyBuf & ".Txt" & vbCrLf
                              AssyBuf = ""
                              Exit Sub
                          End If
                          If Not LOADPSNFILE(AssyBuf) Then
                              Text3.Text = Text3.Text & "--> Unable to load " & AssyBuf & ".Txt" & vbCrLf
                              AssyBuf = ""
                              Exit Sub
                          End If
                          If PSNFileInfo.ScrewStnStatus = "FAIL" Or PSNFileInfo.ScrewStnStatus = "" Then
                              Text3.Text = Text3.Text & "--> Product did not complete previous process" & vbCrLf
                              AssyBuf = ""
                              Exit Sub
                          End If
                          Text3.Text = Text3.Text & "PSN Verified" & vbCrLf
                          PSNFileInfo.FTCheckIn = Today & "," & TimeOfDay

                          If S1_cavitynos = 1 Then 'Cavity#1
                              Text3.Text = Text3.Text & "Current Cavity - C1" & vbCrLf
                              Text100.Text = AssyBuf

                              If Not WRITEPSNFILE("C1") Then
                                  Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                          ElseIf S1_cavitynos = 2 Then  'Cavity#4
                              Text3.Text = Text3.Text & "Current Cavity - C4" & vbCrLf
                              Text103.Text = AssyBuf

                              If Not WRITEPSNFILE("C4") Then
                                  Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                          ElseIf S1_cavitynos = 4 Then  'Cavity#3
                              Text3.Text = Text3.Text & "Current Cavity - C3" & vbCrLf
                              Text102.Text = AssyBuf

                              If Not WRITEPSNFILE("C3") Then
                                  Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                          ElseIf S1_cavitynos = 8 Then  'Cavity#2
                              Text3.Text = Text3.Text & "Current Cavity - C2" & vbCrLf
                              Text101.Text = AssyBuf

                              If Not WRITEPSNFILE("C2") Then
                                  Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                          End If

                          'PSN Verified
                          If Not Modbus.tulisModbus(40104, 1) Then
                              Text3.Text = Text3.Text & "--> Unable to write to PLC - %MW104" & vbCrLf
                              AssyBuf = ""
                              Exit Sub
                          End If
                          lbl_msg.Text = "Please load product and press Start..." & vbCrLf
                          AssyBuf = ""
                      End Sub)
        End If
    End Sub

    '****************************************** SMAC_Comunication ******************************************
    Private Sub SMAC_Comm_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SMAC_Comm.DataReceived
        Smacbuf = SMAC_Comm.ReadExisting()
        If InStr(1, Smacbuf, vbCrLf) <> 0 Then
            Smacbuf = Mid(Smacbuf, 1, InStr(1, Smacbuf, vbCr) - 1)
            Me.Invoke(Sub()
                          Select Case Smacbuf

                              Case "HOMING"

                              Case "HOMING COMPLETED"

                              Case "PGM 1"

                              Case "PGM 1 COMPLETED"

                              Case "PGM 2"

                              Case "PGM 2 COMPLETED"

                              Case "PGM 3"

                              Case "PGM 3 COMPLETED"

                              Case "PGM 4"

                              Case "PGM 4 COMPLETED"

                              Case "PGM 7"

                              Case "PGM 7 COMPLETED"

                              Case "HOMING (2)"

                              Case "HOMING (2) COMPLETED"

                              Case ">"
                                  Text5.Text = Text5.Text & Smacbuf & vbCrLf
                                  Text5.SelectionLength = Len(Text5.Text)
                              Case "PIDS LOADED"
                                  Text5.Text = Text5.Text & Smacbuf & vbCrLf
                                  Text5.SelectionLength = Len(Text5.Text)
                              Case Else
                                  If IsNumeric(Smacbuf) Then
                                      SychResult = CStr(System.Math.Abs(Val(Smacbuf) / 1000))
                                      Text5.Text = Text5.Text & Smacbuf & vbCrLf
                                      Text5.SelectionLength = Len(Text5.Text)

                                  End If

                          End Select
                          Text5.Text = Text5.Text & Smacbuf & vbCrLf
                          Text5.SelectionLength = Len(Text5.Text)
                          Smacbuf = ""
                      End Sub)
        End If
    End Sub

    '****************************************** Chroma_Comunication ******************************************
    Private Sub Chroma_Comm_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles Chroma_Comm.DataReceived
        ChromaBuf = Chroma_Comm.ReadExisting()
        If InStr(1, ChromaBuf, vbCrLf) <> 0 Then
            ChromaBuf = Mid(ChromaBuf, 1, InStr(1, ChromaBuf, vbCrLf) - 1)
            Me.Invoke(Sub()
                          ChromaFBbuf = ChromaBuf
                          ChromaBuf = ""
                      End Sub)
        End If
    End Sub

    Private Function Init_Chroma() As Boolean
        If Not Set_chroma("*IDN?", "Chroma,19053,0,4.30") Then
            Return False
            Exit Function
        End If
        If Not Set_chroma("SYST:LOCK:REQ?", "1") Then
            Return False
            Exit Function
        End If
        Return True
    End Function

    Public Function Set_chroma(Command As String, Reply As String) As Boolean
        Dim Retry As Integer
        Dim Timeout As Double

Retry:
        Timeout = 0
        Chroma_Comm.Write(Command & vbCrLf)
        If Reply <> "" Then
            Do While ChromaFBbuf <> Reply
                Timeout = Timeout + 1
                If Timeout > 60000 Then
                    Retry = Retry + 1
                    If Retry > 4 Then
                        Return False
                        Exit Function
                    Else
                        GoTo Retry
                    End If
                End If
                System.Windows.Forms.Application.DoEvents()
            Loop
        End If
        Return True
    End Function

    Public Function Poll_Chroma_Result() As Boolean
Retry:
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Thread.Sleep(20)
        Chroma_Comm.Write("SOUR:SAFE:RES:ALL?" & vbCrLf)
        Do While ChromaFBbuf = ""
            System.Windows.Forms.Application.DoEvents()

        Loop
        If ChromaFBbuf = "116" Then
            Return True
            Exit Function
        ElseIf ChromaFBbuf = "17" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "18" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "115" Then
            GoTo Retry
        ElseIf ChromaFBbuf = "114" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "113" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "112" Then
            Return False
            Exit Function
        End If
    End Function

    Public Function Set_Chroma_SelfTest() As Boolean
        Dim Timeout As Double
        Dim Retry As Integer
100:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC 100" & vbCrLf)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E+02"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 100
                End If
            End If
        Loop

        Retry = 0
200:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM 0.01" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E-02"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 200
                End If
            End If
        Loop

        Retry = 0
300:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP 0.8" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP?" & vbCrLf)
        Do While ChromaFBbuf <> "+8.000000E-01"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 300
                End If
            End If
        Loop

        Retry = 0
400:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME 1" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E+00"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 400
                End If
            End If
        Loop

        Retry = 0
500:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN(@(1))" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN?" & vbCrLf)
        Do While ChromaFBbuf <> "(@(1))"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 500
                End If
            End If
        Loop

        Retry = 0
600:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN:LOW(@(8))" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN:LOW?" & vbCrLf)
        Do While ChromaFBbuf <> "(@(8))"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 600
                End If
            End If
        Loop
        Return True
        Exit Function
FailComm:
        Return False
    End Function

    Public Function Set_Chroma_HipotTest() As Boolean
        Dim Timeout As Double
        Dim Retry As Integer
100:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC 1920" & vbCrLf)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.920000E+03"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 100
                End If
            End If
        Loop

        Retry = 0
200:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM 0.005" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM?" & vbCrLf)
        Do While ChromaFBbuf <> "+5.000001E-03"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 200
                End If
            End If
        Loop

        Retry = 0
300:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP 0.5" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP?" & vbCrLf)
        Do While ChromaFBbuf <> "+5.000000E-01"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 300
                End If
            End If
        Loop

        Retry = 0
400:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME 1" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E+00"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 400
                End If
            End If
        Loop

        Return True
        Exit Function
FailComm:
        Return False
    End Function

    '****************************************** Sequence Program ******************************************
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim TestResult As Integer
        Dim S1_activity As Integer
        Dim S2_activity As Integer
        Dim S3_activity As Integer
        Dim S4_activity As Integer
        Dim Sys_Interrupt As Integer
        Dim Tagref As String
        Dim Tagnos As String
        Dim TagQty As String
        Dim Tagid As String

        If Button4.Text = "Start" Then
            Tagnos = RD_MULTI_RFID("0000", 10)
            If Tagnos = "NOK" Then GoTo NoChange
            Tagid = RD_MULTI_RFID("0040", 3) 'Read Tag ID
            If Tagnos = "MASTER" Then
                If Tagid = lbl_tagnos.Text Then GoTo NoChange
                If lbl_WOnos.Text <> "MASTER" Then
                    'update the Current WO into the database before changing
                    If CheckWOExist(lbl_WOnos.Text) Then
                        UpdateWO(lbl_WOnos.Text, Label8.Text)
                    Else
                        AddWO(lbl_WOnos.Text)
                        UpdateWO(lbl_WOnos.Text, Label8.Text)
                    End If
                    GoTo WOChange
                ElseIf lbl_WOnos.Text = "MASTER" Then
                    GoTo WOChange
                End If
            ElseIf Tagnos <> lbl_WOnos.Text Then
Master:
                If lbl_WOnos.Text <> "MASTER" Then
                    'update the Current WO into the database before changing
                    If CheckWOExist(lbl_WOnos.Text) Then
                        UpdateWO(lbl_WOnos.Text, Label8.Text)
                    Else
                        AddWO(lbl_WOnos.Text)
                        UpdateWO(lbl_WOnos.Text, Label8.Text)
                    End If
                End If
WOChange:
                Text3.Text = "Changing Series..." & vbCrLf
                Text3.Text = Text3.Text & "Reading info from RFID Tag..." & vbCrLf

                Tagref = RD_MULTI_RFID("0014", 10) 'Read WO Reference from Tag
                If Tagref = "NOK" Then
                    Text3.Text = Text3.Text & "--> Unable to read from RFID Tag" & vbCrLf
                    Text3.Text = Text3.Text & "--> Change Series fail" & vbCrLf
                    Exit Sub
                End If
                TagQty = RD_MULTI_RFID("0028", 10) 'Read WO Qty from Tag
                If TagQty = "NOK" Then
                    Text3.Text = Text3.Text & "--> Unable to read from RFID Tag" & vbCrLf
                    Text3.Text = Text3.Text & "--> Change Series fail" & vbCrLf
                    Exit Sub
                End If
                Tagid = RD_MULTI_RFID("0040", 3) 'Read Tag ID
                If Tagid = "NOK" Then
                    Text3.Text = Text3.Text & "--> Unable to read from RFID Tag" & vbCrLf
                    Text3.Text = Text3.Text & "--> Change Series fail" & vbCrLf
                    Exit Sub
                End If
                Text3.Text = Text3.Text & "loading parameters of new reference..." & vbCrLf
                If Not LoadParameter(Tagref) Then
                    Text3.Text = Text3.Text & "--> Unable to load from database" & vbCrLf
                    Text3.Text = Text3.Text & "--> Change Series fail" & vbCrLf
                    Exit Sub
                End If
                Text3.Text = Text3.Text & "Downloading parameters to PLC..." & vbCrLf
                If Not DownloadParameter() Then
                    Text3.Text = Text3.Text & "--> Unable to download to PLC" & vbCrLf
                    Text3.Text = Text3.Text & "--> Change Series fail" & vbCrLf
                    Exit Sub
                End If
                Label8.Text = "0"
                Label8.ForeColor = Color.Green
                lbl_WOnos.Text = Tagnos
                LoadWOfrRFID.JobNos = Tagnos
                lbl_currentref.Text = Tagref
                LoadWOfrRFID.JobModelName = Tagref
                lbl_wocounter.Text = TagQty
                LoadWOfrRFID.JobQTy = CInt(TagQty)
                lbl_tagnos.Text = Tagid
                LoadWOfrRFID.JobRFIDTag = Tagid
                lbl_ArticleNos.Text = LoadWOfrRFID.JobArticle
                If Tagnos <> "MASTER" Then
                    If CheckWOExist(Tagnos) Then
                        LoadWOfrRFID.JobUnitaryCount = Val(RetrieveWOQty(Tagnos))
                        Label8.Text = CStr(LoadWOfrRFID.JobUnitaryCount)
                    Else
                        AddWO(Tagnos)
                        LoadWOfrRFID.JobUnitaryCount = 0 'Reset output counter
                        Label8.Text = "0"
                    End If
                Else
                    Label8.Text = "0"
                    LoadWOfrRFID.JobUnitaryCount = 0
                End If

                UpdateStnStatus()
                Text3.Text = Text3.Text & "Change Series completed" & vbCrLf
            End If
NoChange:

        End If
yap:
        Ethernet.BackColor = Color.Black
        Sys_Interrupt = Modbus.bacaModbus(40190)
        If Sys_Interrupt > 50 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Light Curtain triggered..."
            frmInterrupt.Button1.Visible = True
            frmInterrupt.Button1.Text = "Continue"
            Exit Sub
        ElseIf Sys_Interrupt = 49 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Light Curtain triggered..."
            frmInterrupt.Button1.Visible = True
            frmInterrupt.Button1.Text = "Continue"
            Exit Sub
        ElseIf Sys_Interrupt = 1 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 2 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 3 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 4 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..."
            frmInterrupt.Button1.Visible = True
            frmInterrupt.Button1.Text = "Reset"
            Exit Sub
        ElseIf Sys_Interrupt = 7 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..." & vbCrLf & "Homing..." & vbCrLf & "Please wait..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 8 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..." & vbCrLf & "Homing..." & vbCrLf & "Please wait..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 9 Then
            frmInterrupt.Show()
            frmInterrupt.Label1.Text = "Emergency Stop..." & vbCrLf & "Homing Completed..."
            frmInterrupt.Button1.Visible = False
            Exit Sub
        ElseIf Sys_Interrupt = 0 Then
            frmInterrupt.Close()
            'frmInterrupt.Hide()
            'Exit Sub
        End If

        'Air Incoming
        If Modbus.bacaModbus(40090) = 1 Then
            Shape30.BackColor = Color.Green
        Else
            Shape30.BackColor = Color.Black
        End If

        'Mode - Run or Stop
        If Modbus.bacaModbus(40091) = 1 Then
            Shape35.BackColor = Color.Green
        Else
            Shape35.BackColor = Color.Black
        End If

        'Indexer Drive Ready
        If Modbus.bacaModbus(40192) = 1 Then
            Shape31.BackColor = Color.Green
        Else
            Shape31.BackColor = Color.Black
        End If

        'Servo On
        If Modbus.bacaModbus(40194) = 1 Then
            Shape33.BackColor = Color.Green
        Else
            Shape33.BackColor = Color.Black
        End If

        'Punching Enable
        If Modbus.bacaModbus(40106) = 0 Then
            Shape34.BackColor = Color.Green
        Else
            Shape34.BackColor = Color.Black
        End If

        'Pin 1 Presence
        If Modbus.bacaModbus(40101) = 0 Then
            Shape36.BackColor = Color.Green
        Else
            Shape36.BackColor = Color.Black
        End If

        'Pin 2 Presence
        If Modbus.bacaModbus(40109) = 0 Then
            frmAlert.Show()
            frmAlert.Label1.Text = "Double Pin feeding..." & vbCrLf & "Please clear pin"
        Else
            frmAlert.Hide()
        End If
        Ethernet.BackColor = Color.Green
        'Clearscreen

        Dim Operation_mode As Integer
        Dim CurrentState As String
        Dim S4_cavity As Integer
        If Button4.Text = "Stop" Then
            If lbl_msg.ForeColor = Color.Yellow Then
                lbl_msg.ForeColor = Color.Black
            Else
                lbl_msg.ForeColor = Color.Yellow
            End If
            Operation_mode = Modbus.bacaModbus(40093)
            If Operation_mode = 0 Then
                Text4.Text = ""
                Text5.Text = ""
                Text6.Text = ""
                Text13.Text = ""
                Text14.Text = ""
                Text11.Text = ""
                Img_Result.Visible = False
                Label79.Text = ""
                Text3.Text = "Operation in progress....."
                lbl_msg.Text = ""
            ElseIf Operation_mode = 2 Then
                If CheckBox1.CheckState = 1 Then
                    lbl_msg.Text = ""
                Else
                    lbl_msg.Text = "Please Scan PSN Label on product....." & vbCrLf & "Load punching pin..." & vbCrLf & "Insert Dummy connector into product and Load..."
                End If
            ElseIf Operation_mode = 3 Then
                lbl_msg.Text = "Press Start to begin punching sequence..."
            ElseIf Operation_mode = 1 Then
                lbl_msg.Text = "Unload product..."
            ElseIf Operation_mode = 4 Then
                lbl_msg.Text = "Load product and Press Start Buttons..."
            ElseIf Operation_mode = 5 Then
                lbl_msg.Text = ""
            End If

            S1_activity = Modbus.bacaModbus(40107)
            Text120.Text = CStr(S1_activity)
            S2_activity = Modbus.bacaModbus(40127)
            Text121.Text = CStr(S2_activity)
            S3_activity = Modbus.bacaModbus(40147)
            Text122.Text = CStr(S3_activity)
            S4_activity = Modbus.bacaModbus(40167)
            Text123.Text = CStr(S4_activity)

            S1_cavitynos = Modbus.bacaModbus(40200)
            If S1_cavitynos = 1 Then
                TestResult = Modbus.bacaModbus(40217)
            ElseIf S1_cavitynos = 2 Then
                TestResult = Modbus.bacaModbus(40247)
            ElseIf S1_cavitynos = 4 Then
                TestResult = Modbus.bacaModbus(40237)
            ElseIf S1_cavitynos = 8 Then
                TestResult = Modbus.bacaModbus(40227)
            End If

            'GoTo Skip
            '------------------- Loading/Punching site activity -------------------
            Select Case S1_activity
                Case 0 'Waiting for user to scan new product
                    ErasePSNFileInfo()
                    If AssyBuf <> "" Then
                        If Len(AssyBuf) < 18 Then
                            AssyBuf = ""
                            Exit Sub
                        End If
                        If Len(AssyBuf) > 18 Then
                            AssyBuf = Microsoft.VisualBasic.Left(AssyBuf, 18)
                        End If

                        If Microsoft.VisualBasic.Left(AssyBuf, 6) <> LoadWOfrRFID.JobArticle Then
                            Text3.Text = "--> PSN - " & AssyBuf & " = wrong reference"
                            AssyBuf = ""
                            Exit Sub
                        Else
                            Text3.Text = "PSN - " & AssyBuf & vbCrLf
                        End If

                        Text3.Text = Text3.Text & "Loading " & AssyBuf & ".Txt..." & vbCrLf
                        If Dir(INIPSNFOLDERPATH & AssyBuf & ".Txt") = "" Then
                            Text3.Text = Text3.Text & "--> Unable to find" & AssyBuf & ".Txt" & vbCrLf
                            AssyBuf = ""
                            Exit Sub
                        End If
                        If Not LOADPSNFILE(AssyBuf) Then
                            Text3.Text = Text3.Text & "--> Unable to load " & AssyBuf & ".Txt" & vbCrLf
                            AssyBuf = ""
                            Exit Sub
                        End If
                        If PSNFileInfo.ScrewStnStatus = "FAIL" Or PSNFileInfo.ScrewStnStatus = "" Then
                            Text3.Text = Text3.Text & "--> Product did not complete previous process" & vbCrLf
                            AssyBuf = ""
                            Exit Sub
                        End If
                        Text3.Text = Text3.Text & "PSN Verified" & vbCrLf
                        PSNFileInfo.FTCheckIn = Today & "," & TimeOfDay

                        If S1_cavitynos = 1 Then 'Cavity#1
                            Text3.Text = Text3.Text & "Current Cavity - C1" & vbCrLf
                            Text100.Text = AssyBuf
                            If Not WRITEPSNFILE("C1") Then
                                Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                AssyBuf = ""
                                Exit Sub
                            End If
                        ElseIf S1_cavitynos = 2 Then  'Cavity#4
                            Text3.Text = Text3.Text & "Current Cavity - C4" & vbCrLf
                            Text103.Text = AssyBuf
                            If Not WRITEPSNFILE("C4") Then
                                Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                AssyBuf = ""
                                Exit Sub
                            End If
                        ElseIf S1_cavitynos = 4 Then  'Cavity#3
                            Text3.Text = Text3.Text & "Current Cavity - C3" & vbCrLf
                            Text102.Text = AssyBuf
                            If Not WRITEPSNFILE("C3") Then
                                Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                AssyBuf = ""
                                Exit Sub
                            End If
                        ElseIf S1_cavitynos = 8 Then  'Cavity#2
                            Text3.Text = Text3.Text & "Current Cavity - C2" & vbCrLf
                            Text101.Text = AssyBuf
                            If Not WRITEPSNFILE("C2") Then
                                Text3.Text = Text3.Text & "--> Unable to register temp file in PSN server" & vbCrLf
                                AssyBuf = ""
                                Exit Sub
                            End If
                        End If
                        'PSN Verified
                        If Not Modbus.tulisModbus(40104, 1) Then
                            Text3.Text = Text3.Text & "--> Unable to write to PLC - %MW104" & vbCrLf
                            AssyBuf = ""
                            Exit Sub
                        End If
                        lbl_msg.Text = "Please load product and press Start..." & vbCrLf
                        AssyBuf = ""
                    End If

                Case 2
                    ErasePSNFileInfo()
                    Select Case S1_cavitynos
                        Case 1
                            Text3.Text = "Retrieving C1 data from Server..." & vbCrLf
                            If Not LOADPSNFILE("C1") Then
                                Text3.Text = "--> Unable to retrieve temp file C1 from PSN folder in server" & vbCrLf
                                If Not Modbus.tulisModbus(40107, 0) Then
                                    Text3.Text = Text3.Text & "--> Unable to reset Site#1" & vbCrLf
                                End If
                                Exit Sub
                            End If
                            Text11.Text = PSNFileInfo.PSN
                            TempPSN = Text11.Text
                            Kill(INIPSNFOLDERPATH & "C1.Txt")
                            Text100.Text = ""
                        Case 2
                            Text3.Text = "Retrieving C4 data from Server..." & vbCrLf
                            If Not LOADPSNFILE("C4") Then
                                Text3.Text = "--> Unable to retrieve temp file C4 from PSN folder in server" & vbCrLf
                                If Not Modbus.tulisModbus(40107, 0) Then
                                    Text3.Text = Text3.Text & "--> Unable to reset Site#1" & vbCrLf
                                End If
                                Exit Sub
                            End If
                            Text11.Text = PSNFileInfo.PSN
                            TempPSN = Text11.Text
                            Kill(INIPSNFOLDERPATH & "C4.Txt")
                            Text103.Text = ""
                        Case 4
                            Text3.Text = "Retrieving C3 data from Server..." & vbCrLf
                            If Not LOADPSNFILE("C3") Then
                                Text3.Text = "--> Unable to retrieve temp file C3 from PSN folder in server" & vbCrLf
                                If Not Modbus.tulisModbus(40107, 0) Then
                                    Text3.Text = Text3.Text & "--> Unable to reset Site#1" & vbCrLf
                                End If
                                Exit Sub
                            End If
                            Text11.Text = PSNFileInfo.PSN
                            TempPSN = Text11.Text
                            Kill(INIPSNFOLDERPATH & "C3.Txt")
                            Text102.Text = ""
                        Case 8
                            Text3.Text = "Retrieving C2 data from Server..." & vbCrLf
                            If Not LOADPSNFILE("C2") Then
                                Text3.Text = "--> Unable to retrieve temp file C2 from PSN folder in server" & vbCrLf
                                If Not Modbus.tulisModbus(40107, 0) Then
                                    Text3.Text = Text3.Text & "--> Unable to reset Site#1" & vbCrLf
                                End If
                                Exit Sub
                            End If
                            Text11.Text = PSNFileInfo.PSN
                            TempPSN = Text11.Text
                            Kill(INIPSNFOLDERPATH & "C2.Txt")
                            Text101.Text = ""
                    End Select
                    EnableCount = False

                    If TestResult = 150 Then 'Pass
                        If PSNFileInfo.FTStatus = "" Or PSNFileInfo.FTStatus = "FAIL" Then
                            EnableCount = True
                        End If
                        Text3.Text = "Punching in progress..." & vbCrLf
                        Text3.Text = Text3.Text & "Hands off...." & vbCrLf
                        'If PSNFileInfo.FTStatus = "PASS" Then
                        '    Modbus.tulisModbus(40106, 1) 'Disable Punch to avoid repeating punch
                        'Else
                        Modbus.tulisModbus(40106, 0) 'Enable punching
                        'End If
                    Else
                        'EnableCount = False
                        PSNFileInfo.FTStatus = "FAIL"
                        Img_Result.Visible = True
                        Img_Result.BackgroundImage = TESE_app.My.Resources.Fail
                        'Text3.Text = "--> " & FailureMode(TestResult) & vbCrLf
                        Label79.Text = TestResult & vbCrLf & FailureMode(TestResult)
                    End If
                    Modbus.tulisModbus(40108, 1) 'To begin punching sequence
                    GoTo 100

                Case 5
                    'Print label if pass
                    If TestResult = 150 Then
                        Text3.Text = Text3.Text & "Printing Schematic & Terminal Labels..."
                        If Not OpenDocument(INILABELTEMPLATEPATH & Parameter.UnitLabelTemplate) Then
                            Text3.Text = Text3.Text & "--> Unable to open label template" & vbCrLf
                            Exit Sub
                        End If
                        If Command3.Text = "Print" Then
                            If EnableCount = True Then
                                PrintLabel(TempPSN)
                            End If
                        End If
                        CloseDocument()
                        If EnableCount = True Then
                            Label8.Text = CStr(Val(Label8.Text) + 1)
                        End If
                        LoadWOfrRFID.JobUnitaryCount = Val(Label8.Text)
                        UpdateStnStatus()
                    Else
                        If Not OpenDocument(INILABELTEMPLATEPATH & "Failure.Lab") Then
                            Text3.Text = Text3.Text & "--> Unable to open label template" & vbCrLf
                            Exit Sub
                        End If
                        If Command3.Text = "Print" Then
                            PrintFailure(TempPSN, Str(TestResult))
                        End If
                        CloseDocument()
                    End If
                    If Not Modbus.tulisModbus(40107, 6) Then
                        Text3.Text = "--> Unable to write to PLC - %MW107" & vbCrLf
                    End If
                    Text3.Text = ""
            End Select

            GoTo skip
100:





        End If
    End Sub
End Class
