Option Explicit On
Public Class frmSelect
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		If TextBox37.Text = "" Then
			MsgBox("Please fill the Reference...")
			Exit Sub
		Else
			If Mid(TextBox37.Text, 1, 6) <> LoadWOfrRFID.JobArticle Then
				MsgBox("Incorrect Reference")
				Exit Sub
			End If
			If Not LOADPSNFILE4REPRINT((TextBox37.Text)) Then
				MsgBox("Unable to load PSN file...")
				Exit Sub
			End If
			If ReprintPSNFileInfo.FTStatus = "PASS" Then
				OpenDocument(INILABELTEMPLATEPATH & Parameter.UnitLabelTemplate)
				ManualPrintLabel()
				CloseCodesoft()
			Else
				MsgBox("Product never pass final test. Unable to print...")
				Exit Sub
			End If
			Me.Close()
		End If
	End Sub

	Public Sub ManualPrintLabel()
		ActiveDoc.Variables.FormVariables.Item("Var14").Value = INILABELIMGPATH & Parameter.UnitLabelPhoto
		ActiveDoc.Variables.FormVariables.Item("Var1").Value = TextBox37.Text & "X13"

		If PrintLabel(1) = False Then
			MsgBox("Error. Can't to print...", MsgBoxStyle.Critical)
		End If
	End Sub
End Class