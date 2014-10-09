Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptScrap
	Inherits System.Windows.Forms.Form
	'Change Log:
	'081203 efj: removed dead variables
	'091803 rjk: Added Unit grid selection when activating this form or selecting fields
	'            and when disactivating return to standard Sector display.
	'            General reformatting. Added setting the initial field on start up.
	'101403 rjk: added dump to pickup additional materials from scrap process
	'101603 rjk: Added set to commands using this form
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace lost with GetLost.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Public strCmd As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'Dim strCmd2 As String
		
		If Option1.Checked Then
			strCmd = strCmd & " ship "
			'    strCmd2 = "sdump *"
		ElseIf Option2.Checked Then 
			strCmd = strCmd & " land "
			'    strCmd2 = "ldump *"
		Else
			strCmd = strCmd & " plane "
			'    strCmd2 = "pdump *"
		End If
		
		strCmd = strCmd & txtUnit.Text
		
		If VB.Left(strCmd, 3) = "set" And Len(txtPrice.Text) > 0 And Len(txtUnit.Text) > 0 Then
			strCmd = strCmd & " " & txtPrice.Text
		End If
		
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'101403 rjk: added dump to pickup additional materials from scrap process
		If InStr(strCmd, "scrap") > 0 Then
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
		End If
		'frmEmpCmd.SubmitEmpireCommand strCmd2, False
		GetLost()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptScrap_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/12/03 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptScrap_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			If txtUnit.Visible Then
				txtUnit.Focus()
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			If txtUnit.Visible Then
				txtUnit.Focus()
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			If txtUnit.Visible Then
				txtUnit.Focus()
			End If
		End If
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Option1.Checked Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		ElseIf Option2.Checked Then 
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		ElseIf Option3.Checked Then 
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End If
	End Sub
End Class