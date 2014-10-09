Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptSectors
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead variables
	'092703 rjk: General reformatting
	'101903 rjk: Modified to support Multiple Sectors Selection using the F4-Conditional Settings
	'            Change to return a range if sectors are somewhat valid otherwise nothing for the range option
	'            Modified to only update the calling form what you have something from the user.
	'
	'102503 rjk: Added support for the multiple territory fields (cbFields)
	'111603 rjk: Changed the name TxtSector to txtSector, code cleanup.
	
	Public HoldSource As System.Windows.Forms.Form
	Public strSectors As String
	
	Private Receiver As System.Windows.Forms.Control
	Private Done As Short
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strParms As String
		Dim strSector As String
		
		On Error Resume Next
		strSector = vbNullString
		
		'Get parms - pull off ? in front
		strParms = txtParm.Text
		If Len(Trim(strParms)) > 0 Then
			If VB.Left(Trim(strParms), 1) = "?" Then
				strParms = Mid(Trim(strParms), 2)
			End If
		End If
		
		'Single Sector
		Dim p1 As Short
		Dim x1 As String
		Dim X2 As String
		Dim y1 As String
		Dim Y2 As String
		If optSect.Checked Then
			strSector = txtSector.Text
			If Len(Trim(txtParm.Text)) > 0 Then
				strSector = strSector & " ?" & strParms
			End If
			
			'Realm
		ElseIf OptRealm.Checked Then 
			strSector = "#" & txtRealm.Text
			If Len(Trim(txtParm.Text)) > 0 Then
				strSector = strSector & " ?" & strParms
			End If
			
			'Range
		ElseIf optRange.Checked Then 
			
			'101903 rjk: Change to return a range if sectors are somewhat valid otherwise nothing
			strSector = ""
			p1 = InStr(txtOrigin.Text, ",")
			If p1 > 0 And Len(txtOrigin.Text) > p1 Then
				x1 = Trim(VB.Left(txtOrigin.Text, p1 - 1))
				y1 = Trim(Mid(txtOrigin.Text, p1 + 1))
				p1 = InStr(txtDest.Text, ",")
				If p1 > 0 And Len(txtDest.Text) > p1 Then
					X2 = Trim(VB.Left(txtDest.Text, p1 - 1))
					Y2 = Trim(Mid(txtDest.Text, p1 + 1))
					strSector = x1 & ":" & X2 & "," & y1 & ":" & Y2
				End If
			End If
			If Len(Trim(txtParm.Text)) > 0 Then
				strSector = strSector & " ?" & strParms
			End If
			'Territory
		ElseIf optTerr.Checked Then 
			'102503 rjk: Added support for the multiple territory fields
			If cbField.SelectedIndex = 0 Then
				strSector = "* ?terr=" & txtTerr.Text
			Else
				strSector = "* ?terr" & CStr(VB6.GetItemData(cbField, cbField.SelectedIndex)) & "=" & txtTerr.Text
			End If
			'strSector = "* ?terr=" + txtTerr.Text
			If Len(Trim(txtParm.Text)) > 0 Then
				strSector = strSector & "&" & strParms
			End If
			'Conditional 101903 rjk: Added
		ElseIf optCond.Checked Then 
			strSector = GetConditionalSectors()
		End If
		
		'Put the string in the Source
		
		If Len(strSector) > 0 Then '101903 rjk: Update the field only you having information
			HoldSource.Enabled = True
			Receiver.Text = strSector
			'Receiver.SetFocus 101903 rjk: Not need to drawn focus to this field as we should be done with it
		End If
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptSectors.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptSectors_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
		
		HoldSource.Enabled = False
		
		'101903 rjk: Added for Multiple Sector Selection using F4 conditional
		If (Receiver.Name = "txtMultOrigin" Or Receiver.Name = "txtMultDest") And ColorScheme = COLORSCHEME_CONDITIONAL Then
			optCond.Enabled = True
		End If
		
		'If the range box is checked - make sure the correct boxes have the focus
		If optRange.Checked And Done = 0 Then
			If Len(txtOrigin.Text) = 0 Or Len(txtDest.Text) > 0 Then
				txtOrigin.Focus()
			Else
				txtDest.Focus()
			End If
			Done = 1
		End If
	End Sub
	
	Private Sub frmPromptSectors_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/12/03 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'Set to recieve focus for clicks on DrawMap
		HoldSource = frmDrawMap.PromptForm
		frmDrawMap.PromptForm = Me
		
		Receiver = HoldSource.ActiveControl
		
		
		'Fill Realm Combo
		Dim X As Short
		For X = 1 To 20
			txtRealm.Items.Add(CStr(X))
		Next X
		txtRealm.SelectedIndex = 0
		optRange.Checked = True
		cbField.SelectedIndex = 0
	End Sub
	
	Private Sub frmPromptSectors_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Done = 0
		frmDrawMap.PromptForm = HoldSource
		HoldSource.Enabled = True
		'UPGRADE_NOTE: Object HoldSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		HoldSource = Nothing
	End Sub
	
	'UPGRADE_WARNING: Event optRange.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optRange_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRange.CheckedChanged
		If eventSender.Checked Then
			On Error Resume Next
			txtOrigin.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event OptRealm.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub OptRealm_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptRealm.CheckedChanged
		If eventSender.Checked Then
			On Error Resume Next
			txtRealm.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optSect.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSect_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSect.CheckedChanged
		If eventSender.Checked Then
			On Error Resume Next
			txtSector.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optTerr.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optTerr_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTerr.CheckedChanged
		If eventSender.Checked Then
			On Error Resume Next
			txtTerr.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtDest.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDest_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDest.TextChanged
		optRange.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		On Error Resume Next
		optRange.Checked = True
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			txtDest.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtRealm.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event txtRealm.Change was upgraded to txtRealm.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub txtRealm_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRealm.TextChanged
		OptRealm.Checked = True
	End Sub
	
	Public Sub SetControls()
		Dim Buffer As String
		Dim parms As String
		Dim x1 As String
		Dim X2 As String
		Dim y1 As String
		Dim Y2 As String
		Dim p1 As Short
		Dim p2 As Short
		Dim p3 As Short
		
		Buffer = Trim(strSectors)
		If Len(Buffer) = 0 Then
			Exit Sub
		End If
		
		'Remove Parameters
		Dim X As Short
		X = InStr(Buffer, "?")
		If X > 0 Then
			parms = Mid(Buffer, X)
			txtParm.Text = Mid(Buffer, X + 1)
			Buffer = VB.Left(Buffer, X - 1)
			
			'Check for territory
			If Mid(parms, 2, 4) = "terr" Then
				p1 = InStr(6, parms, "&")
				If p1 > 0 Then
					txtTerr.Text = Mid(parms, 7, p1 - 7)
					txtParm.Text = Mid(parms, p1 + 1)
				Else
					txtTerr.Text = Mid(parms, 7)
					txtParm.Text = vbNullString
				End If
				Exit Sub
			End If
		End If
		
		'check for Realm
		If VB.Left(Buffer, 1) = "#" Then
			txtRealm.Text = Mid(Buffer, 2)
			Exit Sub
		End If
		
		
		'Check for Range Entry
		p1 = InStr(Buffer, ":")
		If p1 > 0 Then
			x1 = VB.Left(Buffer, p1 - 1)
			p2 = InStr(p1 + 1, Buffer, ",")
			X2 = Mid(Buffer, p1 + 1, p2 - p1 - 1)
			p3 = InStr(p2 + 1, Buffer, ":")
			y1 = Mid(Buffer, p2 + 1, p3 - p2 - 1)
			Y2 = Mid(Buffer, p3 + 1)
			txtOrigin.Text = x1 & "," & y1
			txtDest.Text = X2 & "," & Y2
			Exit Sub
		End If
		
		'Must Just be Single
		txtSector.Text = Buffer
	End Sub
	
	'UPGRADE_WARNING: Event txtRealm.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRealm_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRealm.SelectedIndexChanged
		OptRealm.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event txtSector.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSector_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSector.TextChanged
		If Me.Visible = True Then
			optSect.Checked = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtTerr.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTerr_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTerr.TextChanged
		optTerr.Checked = True
	End Sub
End Class