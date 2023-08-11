Option Explicit On
Imports System.Data.SqlClient
Module modGeneral
    Public Function CheckWOExist(WO As String) As Boolean
        Dim query = "SELECT * FROM TESE.dbo.XCS_13_Tester$ WHERE WONOS = '" & WO & "'"
        Dim dt = KoneksiDB.bacaData(query).Tables(0)

        If dt.Rows(0).Item("WONOS") = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function AddWO(WO As String) As Boolean
        Dim konek As New SqlConnection(Database)

        Dim sql As String = "INSERT INTO TESE.dbo.XCS_13_Tester$ (WONOS,OUTPUT) VALUES('" & WO & "',0)"

        konek.Open()
        Dim sc As New SqlCommand(sql, konek)
        Dim adapter As New SqlDataAdapter(sc)
        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            Return True
            konek.Close()
        Else
            Return False
        End If
    End Function

    Public Function UpdateWO(WO As String, updateqty As String) As Boolean
        Dim konek As New SqlConnection(Database)

        Dim sql As String = "UPDATE TESE.dbo.XCS_13_Tester$ SET OUTPUT='" & updateqty & "' WHERE WONOS='" & WO & "'"

        konek.Open()
        Dim sc As New SqlCommand(sql, konek)
        Dim adapter As New SqlDataAdapter(sc)
        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            Return True
            konek.Close()
        Else
            Return False
        End If
    End Function

    Public Function RetrieveWOQty(WO As String) As String
        Dim readqty As String
        Dim query = "SELECT * FROM TESE.dbo.XCS_13_Tester$ WHERE WONOS = '" & WO & "'"
        Dim dt = KoneksiDB.bacaData(query).Tables(0)

        readqty = dt.Rows(0).Item("OUTPUT")
        Return readqty
    End Function

    'Retrieve the last configuration of tester from file
    Public Sub GetLastConfig()
        Dim Filenum As Integer
        Filenum = FreeFile()
        Dim tempcode As String
        Dim pos1, pos2, pos3, pos4, pos5 As Integer

        FileOpen(Filenum, INISTATUSPATH, OpenMode.Input)
        tempcode = LineInput(Filenum)
        FileClose(Filenum)
        pos1 = InStr(1, tempcode, ",")
        pos2 = InStr(pos1 + 1, tempcode, ",")
        pos3 = InStr(pos2 + 1, tempcode, ",")
        pos4 = InStr(pos3 + 1, tempcode, ",")

        LoadWOfrRFID.JobNos = Mid(tempcode, 1, pos1 - 1)
        LoadWOfrRFID.JobModelName = Mid(tempcode, pos1 + 1, (pos2 - pos1) - 1)
        LoadWOfrRFID.JobQTy = CInt(Mid(tempcode, pos2 + 1, (pos3 - pos2) - 1))
        LoadWOfrRFID.JobRFIDTag = Mid(tempcode, pos3 + 1, (pos4 - pos3) - 1)
        LoadWOfrRFID.JobUnitaryCount = CInt(Mid(tempcode, pos4 + 1))
    End Sub

    Public Function NextCounter(ByVal LastCounter As String) As String
        Dim temp As String

        temp = CStr(CInt(LastCounter) + 1)

        If CInt(temp) >= 9999 Then
            MsgBox("Maximum counter 9999!", MsgBoxStyle.Critical)
            Exit Function
        End If

        If Len(temp) = 1 Then
            temp = "0000" & temp
        ElseIf Len(temp) = 2 Then
            temp = "000" & temp
        ElseIf Len(temp) = 3 Then
            temp = "00" & temp
        ElseIf Len(temp) = 4 Then
            temp = "0" & temp
        End If

        Return temp
    End Function

    Public Sub UpdateStnStatus()
        Dim Filenum1 As Integer

        Filenum1 = FreeFile()
        FileOpen(Filenum1, INISTATUSPATH, OpenMode.Output)
        PrintLine(Filenum1, LoadWOfrRFID.JobNos & "," & LoadWOfrRFID.JobModelName & "," & LoadWOfrRFID.JobQTy & "," & LoadWOfrRFID.JobRFIDTag & "," & LoadWOfrRFID.JobUnitaryCount)
        FileClose(Filenum1)
    End Sub

    Public Sub ErasePSNFileInfo()
        PSNFileInfo.DateCompleted = ""
        PSNFileInfo.DateCreated = ""
        PSNFileInfo.DebugComment = ""
        PSNFileInfo.DebugStatus = ""
        PSNFileInfo.DebugTechnican = ""
        PSNFileInfo.ElectroMagnet = ""
        PSNFileInfo.FTCheckIn = ""
        PSNFileInfo.FTCheckOut = ""
        PSNFileInfo.FTStatus = ""
        PSNFileInfo.MainPCBA = ""
        PSNFileInfo.ModelName = ""
        PSNFileInfo.OperatorID = ""
        PSNFileInfo.PackagingCheckIn = ""
        PSNFileInfo.PackagingCheckOut = ""
        PSNFileInfo.PackagingStatus = ""
        PSNFileInfo.PSN = ""
        PSNFileInfo.RepairDate = ""
        PSNFileInfo.ScrewStnCheckIn = ""
        PSNFileInfo.ScrewStnCheckOut = ""
        PSNFileInfo.ScrewStnStatus = ""
        PSNFileInfo.SecondaryPCBA = ""
        PSNFileInfo.Stn5CheckIn = ""
        PSNFileInfo.Stn5CheckOut = ""
        PSNFileInfo.Stn5Status = ""
        PSNFileInfo.WONos = ""
    End Sub
End Module
