Option Strict Off
Option Explicit On
Friend Class frmPromptImportOffset
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'181006 rjk: Created for setting import offset for intelligence
	
	Public strTelegram As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		strTelegram = vbNullString
		Me.Close()
	End Sub
	
	Private Sub cmdImport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdImport.Click
		Dim iOffsetX, iOffsetY As Short
		If ParseSectors(iOffsetX, iOffsetY, (txtOrigin.Text)) Then
			If Len(strTelegram) > 0 Then
				ImportIntelligenceTelegram(strTelegram, iOffsetX, iOffsetY)
				strTelegram = vbNullString
			Else
				ImportIntelligence(iOffsetX, iOffsetY)
			End If
			Me.Close()
			frmDrawMap.DrawHexes()
		Else
			MsgBox("Invalid Sector", MsgBoxStyle.OKOnly)
		End If
	End Sub
	
	Private Sub frmPromptImportOffset_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptImportOffset_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
End Class