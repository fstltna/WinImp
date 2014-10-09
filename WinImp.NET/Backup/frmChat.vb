Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmChat
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'051904 rjk: Created
	'052204 rjk: Modified to support Flash Chat as well.
	'060604 rjk: Fixed the crash when window was made too small, also fixed the scale for the User combo box.
	'130604 rjk: Changed the initial state for Flash window to country instead all.
	'            Added MsgBox when no Country is selected.
	'190804 rjk: Added reseting the token state for frmEmpCmd when closing the Chat window
	'180206 rjk: Fix the check for empty country list
	'090607 rjk: Fixed the All/Country selection to not change after it has been selected.
	'            Changed the default for Chat.
	'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.
	
	Private cChatColors(5) As Integer
	Private strUsers() As String
	
	'UPGRADE_WARNING: Event cbAllies.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cbAllies.Change was upgraded to cbAllies.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cbAllies_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbAllies.TextChanged
		optUser.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event cbAllies.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbAllies_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbAllies.SelectedIndexChanged
		optUser.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event chkTop.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkTop_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTop.CheckStateChanged
		If chkTop.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		Else
			Call SetWindowPos(Me.Handle.ToInt32, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmChat.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmChat_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If optUser.Checked = False And optAll.Checked = False Then
			If InStr(Me.Text, "Flash") > 0 Then
				optUser.Checked = True
			Else
				optAll.Checked = True
			End If
		End If
	End Sub
	
	Private Sub frmChat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 13 Then
			If Len(txtMessage.Text) > 0 Then
				If InStr(Me.Text, "Flash") > 0 Then
					If optAll.Checked Then
						frmEmpCmd.SubmitEmpireCommand("wall " & txtMessage.Text, False)
						AddText("(" & frmEmpCmd.LoginUser & ")" & txtMessage.Text)
						txtMessage.Text = vbNullString
					ElseIf cbAllies.SelectedIndex <> -1 Then 
						frmEmpCmd.SubmitEmpireCommand("flash " & VB6.GetItemString(cbAllies, cbAllies.SelectedIndex) & " " & txtMessage.Text, False)
						AddText("(" & frmEmpCmd.LoginUser & ")" & txtMessage.Text)
						txtMessage.Text = vbNullString
					Else
						MsgBox("No Country Specified")
					End If
				Else
					If optAll.Checked Then
						frmEmpCmd.SubmitEmpireCommand("lwall " & txtMessage.Text, False)
						txtMessage.Text = vbNullString
					Else
						If Len(txtUser.Text) > 0 Then
							frmEmpCmd.SubmitEmpireCommand("lflash " & txtUser.Text & " " & txtMessage.Text, False)
							AddText("(" & frmEmpCmd.LoginUser & ")" & txtMessage.Text)
							txtMessage.Text = vbNullString
						Else
							MsgBox("No User Specified")
						End If
					End If
				End If
			End If
			KeyAscii = 0
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Public Sub AddText(ByRef strText As String)
		Dim iPosStart As Short
		Dim iPosEnd As Short
		Dim strUser As String
		Dim iColorIndex As Short
		Dim strMessage As String
		Dim bFlash As Boolean
		
		strMessage = strText
		
		rtbText.SelectionStart = Len(rtbText.Text)
		iPosStart = InStr(strMessage, "(")
		iPosEnd = InStr(strMessage, ")")
		If iPosEnd > iPosStart And iPosEnd > 0 Then
			If Mid(strMessage, iPosStart, 2) = "(#" Then
				strUser = Mid(strMessage, iPosStart + 2, iPosEnd - iPosStart - 2) 'country number
				bFlash = True
			Else
				strUser = Mid(strMessage, iPosStart + 1, iPosEnd - iPosStart - 1) 'user name
				bFlash = False
			End If
			If Mid(strMessage, iPosEnd, 2) = "):" Then
				strMessage = VB.Right(strMessage, Len(strMessage) - iPosEnd - 2)
			Else
				strMessage = VB.Right(strMessage, Len(strMessage) - iPosEnd)
			End If
			If Mid(strMessage, 6, 1) = ":" And Mid(strMessage, 9, 1) = ":" Then
				strMessage = LTrim(VB.Right(strMessage, Len(strMessage) - 9))
			End If
			rtbText.SelectionColor = System.Drawing.ColorTranslator.FromOle(GetColor(strUser, bFlash))
			If Len(rtbText.Text) = 0 Then
				rtbText.SelectedText = strMessage
			Else
				rtbText.SelectedText = vbLf & strMessage
			End If
			If chkBeep.CheckState <> System.Windows.Forms.CheckState.Unchecked And frmOptions.bSound Then
				Beep()
			End If
		Else
			rtbText.Text = rtbText.Text & vbLf & "Incorrectly formed message"
		End If
		rtbText.SelectionStart = Len(rtbText.Text)
	End Sub
	
	Private Function GetColor(ByRef strUser As String, ByRef bFlash As Boolean) As Integer
		Dim iIndex As Short
		
		If bFlash And frmOptions.bFlashPlayerColors Then
			'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			GetColor = PlayerColors(Val(strUser) Mod NUMBER_OF_PLAYER_COLORS)
			Exit Function
		End If
		
		For iIndex = 0 To UBound(strUsers)
			If strUsers(iIndex) = strUser Then
				GetColor = cChatColors(iIndex Mod (UBound(cChatColors) + 1))
				Exit Function
			End If
		Next iIndex
		iIndex = UBound(strUsers)
		iIndex = iIndex + 1
		
		ReDim Preserve strUsers(iIndex)
		
		strUsers(iIndex) = strUser
		GetColor = cChatColors(iIndex Mod (UBound(cChatColors) + 1))
	End Function
	
	
	Private Sub frmChat_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ReDim strUsers(0)
		strUsers(0) = frmEmpCmd.LoginUser
		cChatColors(0) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
		cChatColors(1) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
		cChatColors(2) = vbMediumGreen
		cChatColors(3) = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		cChatColors(4) = vbDarkBlue
		cChatColors(5) = vbDarkGreen
	End Sub
	
	'UPGRADE_WARNING: Event frmChat.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmChat_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Exit Sub
		End If
		
		If (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(txtMessage.Height) - VB6.PixelsToTwipsY(cbAllies.Height) - VB6.PixelsToTwipsY(txtMessage.Height) - 850) < 0 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtMessage.Height) + VB6.PixelsToTwipsY(cbAllies.Height) + VB6.PixelsToTwipsY(txtMessage.Height) + 850)
		End If
		rtbText.SetBounds(VB6.TwipsToPixelsX(120), VB6.TwipsToPixelsY(120), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 360), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(txtMessage.Height) - VB6.PixelsToTwipsY(cbAllies.Height) - 850))
		txtMessage.SetBounds(rtbText.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(rtbText.Top) + VB6.PixelsToTwipsY(rtbText.Height) + 120), rtbText.Width, txtMessage.Height)
		txtUser.SetBounds(txtUser.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		cbAllies.SetBounds(cbAllies.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		chkBeep.SetBounds(chkBeep.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		chkTop.SetBounds(chkTop.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		optAll.SetBounds(optAll.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		optUser.SetBounds(optUser.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 850), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
	
	Private Sub frmChat_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		'Only unload if form Draw says OK
		'UPGRADE_ISSUE: Constant vbFormCode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Or UnloadMode = vbFormCode Then
			If (Not frmDrawMap.ShutDown) And Not (InStr(Me.Text, "Flash") > 0 And Not frmOptions.bFlashChat) Then
				Cancel = 1
				Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
				If InStr(Me.Text, "Flash") = 0 Then 'Reset if command gets locked but missing BTU response
					If frmEmpCmd.tokenStatus <> frmEmpCmd.enumToken.tsIdle Then
						frmEmpCmd.tokenStatus = frmEmpCmd.enumToken.tsIdle
					End If
					If frmEmpCmd.CmdinProgress = True Then
						frmEmpCmd.NextCommand()
					End If
				End If
				Exit Sub
			End If
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	Public Sub ControlFlashForm()
		Dim frmFlash As Object
		Dim iFriendlies As Short
		Dim rAlly As Short
		Dim iAllies As Short
		Dim strFriendly() As String
		Dim cnFriendly() As Short
		Dim iIndex As Short
		
		If frmOptions.bFlashChat And Not frmEmpireLogin.bOffline And Not frmDrawMap.ShutDown Then
			cnFriendly = VB6.CopyArray(Nations.GetCountryNumbers())
			If IsArrayEmpty(cnFriendly) < 0 Then
				Exit Sub
			End If
			For iIndex = LBound(cnFriendly) To UBound(cnFriendly)
				If cnFriendly(iIndex) <> CountryNumber Then
					rAlly = Nations.Relation(CountryNumber, cnFriendly(iIndex))
					If rAlly = iREL_ALLIED Or rAlly = iREL_FRIENDLY Then
						ReDim Preserve strFriendly(iFriendlies)
						strFriendly(iFriendlies) = Nations.NationName(cnFriendly(iIndex))
						iFriendlies = iFriendlies + 1
					End If
					If rAlly = iREL_ALLIED Then
						iAllies = iAllies + 1
					End If
				End If
			Next 
			If iAllies = 0 And iFriendlies = 0 Then
				Exit Sub
			End If
			If frmEmpCmd.frmFlash Is Nothing Then
				frmEmpCmd.frmFlash = New frmChat
				frmEmpCmd.frmFlash.Text = "Flash Chat"
				frmEmpCmd.frmFlash.cbAllies.Visible = True
				frmEmpCmd.frmFlash.optUser.Text = "Country"
				frmEmpCmd.frmFlash.txtUser.Visible = False
				frmEmpCmd.frmFlash.Show()
			End If
			frmEmpCmd.frmFlash.cbAllies.Items.Clear()
			For iIndex = 0 To UBound(strFriendly)
				frmEmpCmd.frmFlash.cbAllies.Items.Add(strFriendly(iIndex))
			Next iIndex
			frmEmpCmd.frmFlash.cbAllies.SelectedIndex = -1
			If iAllies = 0 Then
				frmEmpCmd.frmFlash.optAll.Enabled = False
				frmEmpCmd.frmFlash.optUser.Checked = True
			Else
				frmEmpCmd.frmFlash.optAll.Enabled = True
			End If
		Else
			If Not (frmEmpCmd.frmFlash Is Nothing) Then
				frmEmpCmd.frmFlash.Close()
				'UPGRADE_NOTE: Object frmEmpCmd.frmFlash may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				frmEmpCmd.frmFlash = Nothing
			End If
		End If
	End Sub
	
	Public Sub mnuCopyReportWindow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCopyReportWindow.Click
		frmReport.Text = "Telegram/Server Output"
		frmReport.AddReportLine((rtbText.Text))
		frmReport.Show()
	End Sub
	
	Private Sub rtbText_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles rtbText.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		'UPGRADE_ISSUE: Form method frmChat.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		PopupMenu(mnuBaseMenu)
	End Sub
End Class