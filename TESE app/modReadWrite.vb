Option Explicit On
Module modReadWrite
	Public INISTATUSPATH As String 'PATH TO SERVER\FRIDGE
	Public INIPSNFOLDERPATH As String
	Public INILABELTEMPLATEPATH As String
	Public INILABELIMGPATH As String

	Dim FNum As Integer
	Dim LineStr As String

	'Read settings from INI file
	Public Sub ReadINI(Filename As String)
		Dim ItemStr As String
		Dim SectionHeading As String = ""
		Dim pos As Integer

		Dim fullPath As String = System.AppDomain.CurrentDomain.BaseDirectory
		Dim projectFolder As String = fullPath.Replace("\TESE app\bin\Debug\", "").Replace("\TESE app\bin\Release\", "")

		FNum = FreeFile()
		FileOpen(FNum, projectFolder & "\Config\Config.INI", OpenMode.Input)

		Do While Not EOF(FNum)

			LineStr = LineInput(FNum)

			'Check for Section heading
			If Left(LineStr, 1) = "[" Then
				SectionHeading = Mid(LineStr, 2, Len(LineStr) - 2)
			Else
				If InStr(LineStr, "=") > 0 Then
					pos = InStr(LineStr, "=")
					ItemStr = Left(LineStr, pos - 1)

					Select Case UCase(SectionHeading)
						Case "STATUS PATH"
							Select Case UCase(ItemStr)
								Case "PATH" : INISTATUSPATH = Mid(LineStr, pos + 1)
							End Select

						Case "PSN FOLDER" 'LOCAL FILE
							Select Case UCase(ItemStr)
								Case "PATH" : INIPSNFOLDERPATH = Mid(LineStr, pos + 1)
							End Select

						Case "LABEL TEMPLATE PATH" 'LOCAL FILE
							Select Case UCase(ItemStr)
								Case "PATH" : INILABELTEMPLATEPATH = Mid(LineStr, pos + 1)
							End Select

						Case "LABEL PHOTO PATH" 'LOCAL FILE
							Select Case UCase(ItemStr)
								Case "PATH" : INILABELIMGPATH = Mid(LineStr, pos + 1)
							End Select
					End Select
				End If
			End If
		Loop

		FileClose(FNum)
	End Sub

	Public Function LOADPSNFILE(ProductPSN As String) As Boolean
		Dim SectionHeading As String
		Dim pos1, pos2, pos3 As Integer

		FNum = FreeFile()
		If Dir(INIPSNFOLDERPATH & ProductPSN & ".txt") = "" Then
			'SetDefaultINIValues
			'WriteINI
			Return False
			Exit Function
		Else
			FileOpen(FNum, INIPSNFOLDERPATH & ProductPSN & ".txt", OpenMode.Input)

			Do While Not EOF(FNum)

				LineStr = LineInput(FNum)

				'Check for Section heading
				If InStr(LineStr, "[") > 0 And InStr(LineStr, "]") > 0 Then
					pos1 = InStr(LineStr, "[")
					pos2 = InStr(LineStr, "]")
					pos3 = InStr(LineStr, ":")

					SectionHeading = Mid(LineStr, pos1 + 1, pos2 - pos1 - 1)

					Select Case UCase(SectionHeading)
						Case "MODEL"
							PSNFileInfo.ModelName = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE CREATED"
							PSNFileInfo.DateCreated = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE COMPLETED"
							PSNFileInfo.DateCompleted = Trim(Mid(LineStr, pos3 + 1))

						Case "OPERATOR ID"
							PSNFileInfo.OperatorID = Trim(Mid(LineStr, pos3 + 1))

						Case "WORK ORDER NO"
							PSNFileInfo.WONos = Trim(Mid(LineStr, pos3 + 1))

						Case "MAIN PCBA S/N"
							PSNFileInfo.MainPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "SECONDARY PCBA S/N"
							PSNFileInfo.SecondaryPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "ELECTROMAGNET S/N"
							PSNFileInfo.ElectroMagnet = Trim(Mid(LineStr, pos3 + 1))

						Case "PSN"
							PSNFileInfo.PSN = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK IN DATE"
							PSNFileInfo.ScrewStnCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK OUT DATE"
							PSNFileInfo.ScrewStnCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION STATUS"
							PSNFileInfo.ScrewStnStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK IN DATE"
							PSNFileInfo.FTCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK OUT DATE"
							PSNFileInfo.FTCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST SYNCH MEASUREMENT"
							PSNFileInfo.FTSynMeas = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST STATUS"
							PSNFileInfo.FTStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK IN DATE"
							PSNFileInfo.Stn5CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK OUT DATE"
							PSNFileInfo.Stn5CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 STATUS"
							PSNFileInfo.Stn5Status = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK IN DATE"
							PSNFileInfo.VacuumCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK OUT DATE"
							PSNFileInfo.VacummCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM STATUS"
							PSNFileInfo.VacuumStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK IN DATE"
							PSNFileInfo.ConnTestCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK OUT DATE"
							PSNFileInfo.ConnTestCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST STATUS"
							PSNFileInfo.ConnTestStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK IN DATE"
							PSNFileInfo.Vacuum2CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK OUT DATE"
							PSNFileInfo.Vacumm2CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 STATUS"
							PSNFileInfo.Vacuum2Status = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK IN DATE"
							PSNFileInfo.PackagingCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK OUT DATE"
							PSNFileInfo.PackagingCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING STATUS"
							PSNFileInfo.PackagingStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG STATION #10 STATUS"
							PSNFileInfo.DebugStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG COMMENTS"
							PSNFileInfo.DebugComment = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG TECHNICIANS ID"
							PSNFileInfo.DebugTechnican = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG DATE REPAIRED"
							PSNFileInfo.RepairDate = Trim(Mid(LineStr, pos3 + 1))

					End Select
				End If
			Loop

			FileClose(FNum)
			Return True
		End If
	End Function

	'Write data to PSN file
	Public Function WRITEPSNFILE(ProductPSN As String) As Boolean
		On Error GoTo ErrorHandler
		FNum = FreeFile()
		FileOpen(FNum, INIPSNFOLDERPATH & ProductPSN & ".txt", OpenMode.Output)

		PrintLine(FNum)
		PrintLine(FNum, "[MODEL] : " & PSNFileInfo.ModelName)
		PrintLine(FNum)
		PrintLine(FNum, "[DATE CREATED] : " & PSNFileInfo.DateCreated)
		PrintLine(FNum)
		PrintLine(FNum, "[DATE COMPLETED] : " & PSNFileInfo.DateCompleted)
		PrintLine(FNum)
		PrintLine(FNum, "[OPERATOR ID] : " & PSNFileInfo.OperatorID)
		PrintLine(FNum)
		PrintLine(FNum, "[WORK ORDER NO] : " & PSNFileInfo.WONos)
		PrintLine(FNum)
		PrintLine(FNum, "[MAIN PCBA S/N] : " & PSNFileInfo.MainPCBA)
		PrintLine(FNum)
		PrintLine(FNum, "[SECONDARY PCBA S/N] : " & PSNFileInfo.SecondaryPCBA)
		PrintLine(FNum)
		PrintLine(FNum, "[ELECTROMAGNET S/N] : " & PSNFileInfo.ElectroMagnet)
		PrintLine(FNum)
		PrintLine(FNum, "[PSN] : " & PSNFileInfo.PSN)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION CHECK IN DATE] : " & PSNFileInfo.ScrewStnCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION CHECK OUT DATE] : " & PSNFileInfo.ScrewStnCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION STATUS] : " & PSNFileInfo.ScrewStnStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST CHECK IN DATE] : " & PSNFileInfo.FTCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST CHECK OUT DATE] : " & PSNFileInfo.FTCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST SYNCH MEASUREMENT] : " & PSNFileInfo.FTSynMeas)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST STATUS] : " & PSNFileInfo.FTStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 CHECK IN DATE] : " & PSNFileInfo.Stn5CheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 CHECK OUT DATE] : " & PSNFileInfo.Stn5CheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 STATUS] : " & PSNFileInfo.Stn5Status)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM CHECK IN DATE] : " & PSNFileInfo.VacuumCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM CHECK OUT DATE] : " & PSNFileInfo.VacummCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM STATUS] : " & PSNFileInfo.VacuumStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST CHECK IN DATE] : " & PSNFileInfo.ConnTestCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST CHECK OUT DATE] : " & PSNFileInfo.ConnTestCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST STATUS] : " & PSNFileInfo.ConnTestStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 CHECK IN DATE] : " & PSNFileInfo.Vacuum2CheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 CHECK OUT DATE] : " & PSNFileInfo.Vacumm2CheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 STATUS] : " & PSNFileInfo.Vacuum2Status)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING CHECK IN DATE] : " & PSNFileInfo.PackagingCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING CHECK OUT DATE] : " & PSNFileInfo.PackagingCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING STATUS] : " & PSNFileInfo.PackagingStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG STATION #10 STATUS] : " & PSNFileInfo.DebugStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG COMMENTS] : " & PSNFileInfo.DebugComment)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG TECHNICIANS ID] : " & PSNFileInfo.DebugTechnican)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG DATE REPAIRED] : " & PSNFileInfo.RepairDate)
		FileClose(FNum)

		Return True
		Exit Function
ErrorHandler:
		Return False
	End Function

	Public Function LOADPSNFILE4REPRINT(ProductPSN As String) As Boolean
		Dim SectionHeading As String
		Dim pos1, pos2, pos3 As Integer

		FNum = FreeFile()
		If Dir(INIPSNFOLDERPATH & ProductPSN & ".txt") = "" Then
			'SetDefaultINIValues
			'WriteINI
			Return False
			Exit Function
		Else
			FileOpen(FNum, INIPSNFOLDERPATH & ProductPSN & ".txt", OpenMode.Input)

			Do While Not EOF(FNum)

				LineStr = LineInput(FNum)

				'Check for Section heading
				If InStr(LineStr, "[") > 0 And InStr(LineStr, "]") > 0 Then
					pos1 = InStr(LineStr, "[")
					pos2 = InStr(LineStr, "]")
					pos3 = InStr(LineStr, ":")

					SectionHeading = Mid(LineStr, pos1 + 1, pos2 - pos1 - 1)

					Select Case UCase(SectionHeading)
						Case "MODEL"
							ReprintPSNFileInfo.ModelName = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE CREATED"
							ReprintPSNFileInfo.DateCreated = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE COMPLETED"
							ReprintPSNFileInfo.DateCompleted = Trim(Mid(LineStr, pos3 + 1))

						Case "OPERATOR ID"
							ReprintPSNFileInfo.OperatorID = Trim(Mid(LineStr, pos3 + 1))

						Case "WORK ORDER NO"
							ReprintPSNFileInfo.WONos = Trim(Mid(LineStr, pos3 + 1))

						Case "MAIN PCBA S/N"
							ReprintPSNFileInfo.MainPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "SECONDARY PCBA S/N"
							ReprintPSNFileInfo.SecondaryPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "ELECTROMAGNET S/N"
							ReprintPSNFileInfo.ElectroMagnet = Trim(Mid(LineStr, pos3 + 1))

						Case "PSN"
							ReprintPSNFileInfo.PSN = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK IN DATE"
							ReprintPSNFileInfo.ScrewStnCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK OUT DATE"
							ReprintPSNFileInfo.ScrewStnCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION STATUS"
							ReprintPSNFileInfo.ScrewStnStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK IN DATE"
							ReprintPSNFileInfo.FTCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK OUT DATE"
							ReprintPSNFileInfo.FTCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST SYNCH MEASUREMENT"
							ReprintPSNFileInfo.FTSynMeas = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST STATUS"
							ReprintPSNFileInfo.FTStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK IN DATE"
							ReprintPSNFileInfo.Stn5CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK OUT DATE"
							ReprintPSNFileInfo.Stn5CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 STATUS"
							ReprintPSNFileInfo.Stn5Status = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK IN DATE"
							ReprintPSNFileInfo.VacuumCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK OUT DATE"
							ReprintPSNFileInfo.VacummCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM STATUS"
							ReprintPSNFileInfo.VacuumStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK IN DATE"
							ReprintPSNFileInfo.ConnTestCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK OUT DATE"
							ReprintPSNFileInfo.ConnTestCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST STATUS"
							ReprintPSNFileInfo.ConnTestStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK IN DATE"
							ReprintPSNFileInfo.Vacuum2CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK OUT DATE"
							ReprintPSNFileInfo.Vacumm2CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 STATUS"
							ReprintPSNFileInfo.Vacuum2Status = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK IN DATE"
							ReprintPSNFileInfo.PackagingCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK OUT DATE"
							ReprintPSNFileInfo.PackagingCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING STATUS"
							ReprintPSNFileInfo.PackagingStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG STATION #10 STATUS"
							ReprintPSNFileInfo.DebugStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG COMMENTS"
							ReprintPSNFileInfo.DebugComment = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG TECHNICIANS ID"
							ReprintPSNFileInfo.DebugTechnican = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG DATE REPAIRED"
							ReprintPSNFileInfo.RepairDate = Trim(Mid(LineStr, pos3 + 1))

					End Select
				End If
			Loop

			FileClose(FNum)
			Return True
		End If
	End Function

	Public Function LOADPSNFILE4S4(cavityno As String) As Boolean
		Dim SectionHeading As String
		Dim pos1, pos2, pos3 As Integer

		FNum = FreeFile()
		If Dir(INIPSNFOLDERPATH & cavityno & ".txt") = "" Then
			'SetDefaultINIValues
			'WriteINI
			Return False
			Exit Function
		Else
			FileOpen(FNum, INIPSNFOLDERPATH & cavityno & ".txt", OpenMode.Input)

			Do While Not EOF(FNum)

				LineStr = LineInput(FNum)

				'Check for Section heading
				If InStr(LineStr, "[") > 0 And InStr(LineStr, "]") > 0 Then
					pos1 = InStr(LineStr, "[")
					pos2 = InStr(LineStr, "]")
					pos3 = InStr(LineStr, ":")

					SectionHeading = Mid(LineStr, pos1 + 1, pos2 - pos1 - 1)

					Select Case UCase(SectionHeading)
						Case "MODEL"
							S4PSNFileinfo.ModelName = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE CREATED"
							S4PSNFileinfo.DateCreated = Trim(Mid(LineStr, pos3 + 1))

						Case "DATE COMPLETED"
							S4PSNFileinfo.DateCompleted = Trim(Mid(LineStr, pos3 + 1))

						Case "OPERATOR ID"
							S4PSNFileinfo.OperatorID = Trim(Mid(LineStr, pos3 + 1))

						Case "WORK ORDER NO"
							S4PSNFileinfo.WONos = Trim(Mid(LineStr, pos3 + 1))

						Case "MAIN PCBA S/N"
							S4PSNFileinfo.MainPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "SECONDARY PCBA S/N"
							S4PSNFileinfo.SecondaryPCBA = Trim(Mid(LineStr, pos3 + 1))

						Case "ELECTROMAGNET S/N"
							S4PSNFileinfo.ElectroMagnet = Trim(Mid(LineStr, pos3 + 1))

						Case "PSN"
							S4PSNFileinfo.PSN = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK IN DATE"
							S4PSNFileinfo.ScrewStnCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION CHECK OUT DATE"
							S4PSNFileinfo.ScrewStnCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "SCREWING STATION STATUS"
							S4PSNFileinfo.ScrewStnStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK IN DATE"
							S4PSNFileinfo.FTCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST CHECK OUT DATE"
							S4PSNFileinfo.FTCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST SYNCH MEASUREMENT"
							S4PSNFileinfo.FTSynMeas = Trim(Mid(LineStr, pos3 + 1))

						Case "FINAL TEST STATUS"
							S4PSNFileinfo.FTStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK IN DATE"
							S4PSNFileinfo.Stn5CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 CHECK OUT DATE"
							S4PSNFileinfo.Stn5CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "STATION 5 STATUS"
							S4PSNFileinfo.Stn5Status = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK IN DATE"
							S4PSNFileinfo.VacuumCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM CHECK OUT DATE"
							S4PSNFileinfo.VacummCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM STATUS"
							S4PSNFileinfo.VacuumStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK IN DATE"
							S4PSNFileinfo.ConnTestCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST CHECK OUT DATE"
							S4PSNFileinfo.ConnTestCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "CONNECTOR TEST STATUS"
							S4PSNFileinfo.ConnTestStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK IN DATE"
							S4PSNFileinfo.Vacuum2CheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 CHECK OUT DATE"
							S4PSNFileinfo.Vacumm2CheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "VACUUM #2 STATUS"
							S4PSNFileinfo.Vacuum2Status = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK IN DATE"
							S4PSNFileinfo.PackagingCheckIn = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING CHECK OUT DATE"
							S4PSNFileinfo.PackagingCheckOut = Trim(Mid(LineStr, pos3 + 1))

						Case "PACKAGING STATUS"
							S4PSNFileinfo.PackagingStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG STATION #10 STATUS"
							S4PSNFileinfo.DebugStatus = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG COMMENTS"
							S4PSNFileinfo.DebugComment = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG TECHNICIANS ID"
							S4PSNFileinfo.DebugTechnican = Trim(Mid(LineStr, pos3 + 1))

						Case "DEBUG DATE REPAIRED"
							S4PSNFileinfo.RepairDate = Trim(Mid(LineStr, pos3 + 1))

					End Select
				End If
			Loop
			FileClose(FNum)
			Return True
		End If
	End Function

	'Write data to PSN file
	Public Function WRITEPSNFILE4S4(cavityno As String) As Boolean
		On Error GoTo ErrorHandler
		FNum = FreeFile()
		FileOpen(FNum, INIPSNFOLDERPATH & cavityno & ".txt", OpenMode.Output)

		PrintLine(FNum)
		PrintLine(FNum, "[MODEL] : " & S4PSNFileinfo.ModelName)
		PrintLine(FNum)
		PrintLine(FNum, "[DATE CREATED] : " & S4PSNFileinfo.DateCreated)
		PrintLine(FNum)
		PrintLine(FNum, "[DATE COMPLETED] : " & S4PSNFileinfo.DateCompleted)
		PrintLine(FNum)
		PrintLine(FNum, "[OPERATOR ID] : " & S4PSNFileinfo.OperatorID)
		PrintLine(FNum)
		PrintLine(FNum, "[WORK ORDER NO] : " & S4PSNFileinfo.WONos)
		PrintLine(FNum)
		PrintLine(FNum, "[MAIN PCBA S/N] : " & S4PSNFileinfo.MainPCBA)
		PrintLine(FNum)
		PrintLine(FNum, "[SECONDARY PCBA S/N] : " & S4PSNFileinfo.SecondaryPCBA)
		PrintLine(FNum)
		PrintLine(FNum, "[ELECTROMAGNET S/N] : " & S4PSNFileinfo.ElectroMagnet)
		PrintLine(FNum)
		PrintLine(FNum, "[PSN] : " & S4PSNFileinfo.PSN)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION CHECK IN DATE] : " & S4PSNFileinfo.ScrewStnCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION CHECK OUT DATE] : " & S4PSNFileinfo.ScrewStnCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[SCREWING STATION STATUS] : " & S4PSNFileinfo.ScrewStnStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST CHECK IN DATE] : " & S4PSNFileinfo.FTCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST CHECK OUT DATE] : " & S4PSNFileinfo.FTCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST SYNCH MEASUREMENT] : " & S4PSNFileinfo.FTSynMeas)
		PrintLine(FNum)
		PrintLine(FNum, "[FINAL TEST STATUS] : " & S4PSNFileinfo.FTStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 CHECK IN DATE] : " & S4PSNFileinfo.Stn5CheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 CHECK OUT DATE] : " & S4PSNFileinfo.Stn5CheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[STATION 5 STATUS] : " & S4PSNFileinfo.Stn5Status)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM CHECK IN DATE] : " & S4PSNFileinfo.VacuumCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM CHECK OUT DATE] : " & S4PSNFileinfo.VacummCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM STATUS] : " & S4PSNFileinfo.VacuumStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST CHECK IN DATE] : " & S4PSNFileinfo.ConnTestCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST CHECK OUT DATE] : " & S4PSNFileinfo.ConnTestCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[CONNECTOR TEST STATUS] : " & S4PSNFileinfo.ConnTestStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 CHECK IN DATE] : " & S4PSNFileinfo.Vacuum2CheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 CHECK OUT DATE] : " & S4PSNFileinfo.Vacumm2CheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[VACUUM #2 STATUS] : " & S4PSNFileinfo.Vacuum2Status)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING CHECK IN DATE] : " & S4PSNFileinfo.PackagingCheckIn)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING CHECK OUT DATE] : " & S4PSNFileinfo.PackagingCheckOut)
		PrintLine(FNum)
		PrintLine(FNum, "[PACKAGING STATUS] : " & S4PSNFileinfo.PackagingStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG STATION #10 STATUS] : " & S4PSNFileinfo.DebugStatus)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG COMMENTS] : " & S4PSNFileinfo.DebugComment)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG TECHNICIANS ID] : " & S4PSNFileinfo.DebugTechnican)
		PrintLine(FNum)
		PrintLine(FNum, "[DEBUG DATE REPAIRED] : " & S4PSNFileinfo.RepairDate)
		FileClose(FNum)

		Return True
		Exit Function
ErrorHandler:
		Return False
	End Function
End Module
