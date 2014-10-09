Option Strict Off
Option Explicit On
Friend Class frmPromptShoot
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Added Option Explicit and removed dead variables
	'092703 rjk: Set the quantity field based on the option and
	'            origin.  General Reformatting.
	'093003 rjk: only switch to txtNum field if origin is valid.
	'093003 rjk: do not set txtNum for multiple sector locations.
	'093003 rjk: Set the options based on the sector location.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'120203 rjk: Changed to make civ the default if present.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		If Option1.Checked Then
			strCmd = "u "
		Else
			strCmd = "c "
		End If
		
		strCmd = "shoot " & strCmd & txtOrigin.Text & " " & txtNum.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'frmEmpCmd.SubmitEmpireCommand "dump " + txtOrigin.Text, False
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptShoot.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptShoot_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		SetTextNum()
	End Sub
	
	Private Sub frmPromptShoot_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptShoot_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			SetTextNum()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			SetTextNum()
		End If
	End Sub
	
	Private Sub txtOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.DoubleClick
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptSectors)
		frmPromptSectors.strSectors = txtOrigin.Text
		frmPromptSectors.SetControls()
		frmPromptSectors.Text = "Select Sectors"
		frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
		frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
		frmPromptSectors.Show()
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		'093003 rjk: Set the options based on the sector location.
		Dim secx As Short
		Dim secy As Short
		
		If Len(txtOrigin.Text) > 0 Then
			If ParseSectors(secx, secy, (txtOrigin.Text)) Then
				rsSectors.Seek("=", secx, secy)
				If Not rsSectors.NoMatch Then
					'120203 rjk: Switched the order to make civ the default if present.
					If CDbl(CStr(rsSectors.Fields("uw").Value)) > 0 Then
						Option1.Enabled = True
						Option1.Checked = True
					Else
						Option1.Enabled = False
					End If
					If CDbl(CStr(rsSectors.Fields("civ").Value)) > 0 Then
						Option2.Enabled = True
						Option2.Checked = True
					Else
						Option2.Enabled = False
					End If
				Else
					Option1.Enabled = False
					Option2.Enabled = False
				End If
			Else
				Option1.Enabled = True
				Option2.Enabled = True
				txtNum.Text = vbNullString
			End If
		Else
			Option1.Enabled = False
			Option2.Enabled = False
		End If
		
		SetTextNum()
	End Sub
	
	Private Sub SetTextNum()
		Dim secx As Short
		Dim secy As Short
		
		If Len(txtOrigin.Text) > 0 Then
			If ParseSectors(secx, secy, (txtOrigin.Text)) Then
				rsSectors.Seek("=", secx, secy)
				If Not rsSectors.NoMatch Then
					If Option2.Checked Then
						txtNum.Text = CStr(rsSectors.Fields("civ").Value)
					ElseIf Option1.Checked Then 
						txtNum.Text = CStr(rsSectors.Fields("uw").Value)
					Else
						txtNum.Text = vbNullString
					End If
					If txtNum.Visible Then
						txtNum.Focus()
					End If
				Else
					txtNum.Text = vbNullString
					If txtOrigin.Visible Then '093003 rjk: only switch to txtNum field if origin is valid.
						txtOrigin.Focus()
					End If
				End If
				'Else multiple sector selection, do not erase txtNum
			End If
		Else
			txtNum.Text = vbNullString
			If txtOrigin.Visible Then '093003 rjk: only switch to txtNum field if origin is valid.
				txtOrigin.Focus()
			End If
		End If
	End Sub
End Class