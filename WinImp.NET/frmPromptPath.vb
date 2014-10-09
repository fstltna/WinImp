Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmPromptPath
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead variables
	'092703 rjk: General reformatting
	'101903 rjk: Added Condition Button for Multiple Sector Selection
	'111403 rjk: Fixed so the keyboard can be used to select direction
	'            Added Initial Field selection and focus changes on the return/home/reset buttons.
	
	Public HoldSource As System.Windows.Forms.Form
	' Public strSectors As String    8/2003 efj  removed
	
	' Private ActivateUnitBox As Boolean    8/2003 efj  removed
	Private Receiver As System.Windows.Forms.Control
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.FinishPath()
		Me.Close()
	End Sub
	
	'101903 rjk: Return all the sectors that are conditional selected
	Private Sub cmdConditional_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConditional.Click
		frmDrawMap.FinishPath()
		Receiver.Text = GetConditionalSectors
		Me.Close()
	End Sub
	
	Private Sub cmdHome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHome.Click
		Dim x1 As Short
		Dim X2 As Short
		Dim y1 As Short
		Dim Y2 As Short
		Dim i As Short
		Dim strPath As String
		
		'Compute the path home, and add to path string
		If ParseSectors(x1, y1, (lblEndSector.Text)) Then
			If ParseSectors(X2, Y2, (lblSector.Text)) Then
				strPath = EmpirePathDirection(X2 - x1, Y2 - y1)
				For i = 1 To Len(strPath)
					frmDrawMap.ExtendPath(Mid(strPath, i, 1))
				Next i
			End If
		End If
		cmdOK.Focus() '111403 rjk: Draw the focus the Okay button to finish form.
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		frmDrawMap.FinishPath()
		Receiver.Text = txtPath.Text & "h"
		Me.Close()
	End Sub
	
	Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
		frmDrawMap.FinishPath()
		frmDrawMap.StartPath((lblSector.Text))
		lblEndSector.Text = lblSector.Text
		txtPath.Text = vbNullString
		txtPath.Focus() '111403 rjk: Draw the focus the txtPath so the keyboard works
	End Sub
	
	Private Sub cmdReturn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReturn.Click
		Dim i As Short
		Dim strPath As String
		Dim ch As String
		
		'Compute the path home, and add to path string
		strPath = txtPath.Text
		For i = Len(strPath) To 1 Step -1
			Select Case Mid(strPath, i, 1)
				Case "y"
					ch = "n"
				Case "u"
					ch = "b"
				Case "j"
					ch = "g"
				Case "n"
					ch = "y"
				Case "b"
					ch = "u"
				Case "g"
					ch = "j"
				Case Else
					ch = vbNullString
			End Select
			frmDrawMap.ExtendPath(ch)
		Next i
		cmdOK.Focus() '111403 rjk: Draw the focus the Okay button to finish form.
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptPath.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptPath_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		HoldSource.Enabled = False
		If Not frmDrawMap.DrawingPath Then
			frmDrawMap.StartPath((lblSector.Text))
		End If
		'101903 rjk: Only enable the button if the path is empty, it is multiple sector field, and conditional is active
		If Len(txtPath.Text) = 0 And Receiver.Name = "txtMultPath" And ColorScheme = COLORSCHEME_CONDITIONAL Then
			cmdConditional.Visible = True
		Else
			cmdConditional.Visible = False
		End If
		txtPath.Focus() '111403 rjk: Draw the focus the txtPath so the keyboard works
	End Sub
	
	Private Sub frmPromptPath_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'Set to recieve focus for clicks on DrawMap
		HoldSource = frmDrawMap.PromptForm
		frmDrawMap.PromptForm = Me
		
		Receiver = HoldSource.ActiveControl
	End Sub
	
	Private Sub frmPromptPath_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmDrawMap.PromptForm = HoldSource
		HoldSource.Enabled = True
		On Error Resume Next
		HoldSource.Activate()
		
		'UPGRADE_NOTE: Object HoldSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		HoldSource = Nothing
	End Sub
	
	'UPGRADE_WARNING: Event txtPath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.TextChanged
		'Update path lenght
		Label4.Text = CStr(Len(txtPath.Text))
		
		'update the path cost
		Dim c As Single
		c = PathCost((Me.lblSector.Text), (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
		If c > 0 Then
			Label6.Text = "Path Cost"
			Label7.Text = CStr(CSng(CInt(c * 1000) / 1000))
		Else
			Label6.Text = vbNullString
			Label7.Text = vbNullString
		End If
		'101903 rjk: Only enable the button if the path is empty, it is multiple sector field, and conditional is active
		If Len(txtPath.Text) = 0 And Receiver.Name = "txtMultPath" And ColorScheme = COLORSCHEME_CONDITIONAL Then
			cmdConditional.Visible = True
		Else
			cmdConditional.Visible = False
		End If
	End Sub
	
	Private Sub txtPath_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPath.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim ch As String
		Dim x As Short
		
		'If it is a valid direction entry, then add to path
		ch = Chr(KeyAscii)
		x = InStr("gyujbn", ch) '111403 rjk: Reverse order to make the check work.
		If x > 0 Then
			frmDrawMap.ExtendPath(ch)
			KeyAscii = 0
			GoTo EventExitSub
		End If
		
		'If its a backspace, then remove the last path entry
		If KeyAscii = System.Windows.Forms.Keys.Back Then
			frmDrawMap.RemoveFromPath()
			KeyAscii = 0
			GoTo EventExitSub
		End If
		
		KeyAscii = 0
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class