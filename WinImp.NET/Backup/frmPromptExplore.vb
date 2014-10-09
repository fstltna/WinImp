Option Strict Off
Option Explicit On
Friend Class frmPromptExplore
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: removed dead variables
	'090103 rjk: Added for additional exploration modes
	'092803 rjk: General reformatting.
	'101503 rjk: Disable the mode when in path mode
	'101703 rjk: Added Strength fields to Sector database
	'110703 rjk: Fixed a problem filling the path from the frmPromptPath form.
	'112003 rjk: Added option to control strength updates
	'122903 rjk: Fixed help to use the Title label
	'270104 rjk: When selecting population, do not change the path option to radius
	'050204 rjk: Select the radius option when the radius combo is selected.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'010804 rjk: Select the TakePath option when returning from the Path Selection form.
	'250804 rjk: Disabled Civilians for 2K4.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim strUnit As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label1.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strnum As String
		Dim exploreMode As WinAceRoutines.enumExploreMode '090103 rjk: Added mode for user to select exploration pattern
		
		strnum = CStr(Val(Combo1.Text))
		If Option3.Checked Then
			
			'Compute text Destination
			strCmd = "explore " & strUnit & " " & txtOrigin.Text & " " & strnum & " " & txtPath.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			'database update
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		Else
			If Mode(1).Checked = True Then
				exploreMode = WinAceRoutines.enumExploreMode.EXP_COMPLETE
			ElseIf Mode(2).Checked = True Then 
				exploreMode = WinAceRoutines.enumExploreMode.EXP_BORDER
			ElseIf Mode(3).Checked = True Then 
				exploreMode = WinAceRoutines.enumExploreMode.EXP_WHEEL
			Else 'should not happen
				exploreMode = WinAceRoutines.enumExploreMode.EXP_COMPLETE
			End If
			ExploreOut((txtOrigin.Text), strUnit, CShort(Combo2.Text), 1, 0, Val(Combo1.Text), exploreMode)
		End If
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event Combo1.Change was upgraded to Combo1.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub Combo1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.TextChanged
		txtOrigin.Focus()
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		txtOrigin.Focus()
	End Sub
	
	'UPGRADE_WARNING: Event Combo2.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.SelectedIndexChanged
		Option4.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event Combo2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event Combo2.Change was upgraded to Combo2.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub Combo2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.TextChanged
		Option4.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptExplore.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptExplore_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtOrigin.Focus()
		
		Dim X As Short
		Dim sx As Short
		Dim sy As Short
		
		Combo2.Items.Clear()
		For X = 1 To MaxExploreRadius
			Combo2.Items.Add(CStr(X))
		Next X
		
		Combo1.SelectedIndex = 0
		Combo2.SelectedIndex = Combo2.Items.Count - 1
		
		'if there are no civs in the sector, default to mil
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				If rsSectors.Fields("civ").Value < 2 Then
					Option2.Checked = True
				End If
			End If
		End If
		If frmOptions.b2K4 Then
			Option1.Enabled = False
		End If
		If Len(txtPath.Text) > 0 Then
			Option3.Checked = True
		End If
	End Sub
	
	Private Sub frmPromptExplore_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long   removed efj 8/2003
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		Option1.Checked = True
		strUnit = "c"
	End Sub
	
	Private Sub frmPromptExplore_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			strUnit = "c"
			txtOrigin.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			strUnit = "m"
			txtOrigin.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			On Error Resume Next '101503 rjk: Removed not needed?
			'110703 rjk: Added back. Used when filling txtPath the form the frmPromptPath form,
			'            it is not the active form and can not accept the focus.
			txtPath.Focus()
			'101503 rjk: Disable the Explore pattern frame, not appropriate for this mode
			frameMode.Visible = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option4.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option4_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option4.CheckedChanged
		If eventSender.Checked Then
			'101503 rjk: Ensure the Explore pattern frame is displayed
			frameMode.Visible = True
		End If
	End Sub
	
	Private Sub txtPath_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.Click
		Option3.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event txtPath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.TextChanged
		Option3.Checked = True
	End Sub
	
	Private Sub txtPath_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.DoubleClick
		'Make sure starting sector is valid before prompting
		Dim sx As Short
		Dim sy As Short
		If Not ParseSectors(sx, sy, (txtOrigin.Text)) Then
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptPath)
		frmPromptPath.lblSector.Text = txtOrigin.Text
		frmPromptPath.lblEndSector.Text = txtOrigin.Text
		frmPromptPath.Text = "Exploration Route"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
	End Sub
End Class