Option Strict Off
Option Explicit On
Friend Class frmPromptTorp
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'091803 rjk: Added Unit grid selection when activating this form or selecting fields
	'            and when disactivating return to standard Sector display.
	'            General reformatting. Added setting the initial field on start up.
	'180206 rjk: Replace sdump with GetShips.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim num As Short
		Dim X As Short
		
		strCmd = "torpedo " & txtUnit.Text & " " & txtTarget.Text
		
		'Execute fire command
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			num = Val(Combo1.Text)
		Else
			num = 1
		End If
		
		If num > 1 Then
			frmEmpCmd.SubmitEmpireCommand("bf1", False)
			For X = 1 To num
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			Next 
			frmEmpCmd.SubmitEmpireCommand("bf2", False)
		Else
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetShips()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event Combo1.Change was upgraded to Combo1.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub Combo1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.TextChanged
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub frmPromptTorp_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		txtUnit.Visible = True
	End Sub
	
	Private Sub frmPromptTorp_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_S_TORP, True, True, False)
	End Sub
	
	Private Sub txtTarget_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTarget.Enter
		frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udENEMY, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
	End Sub
End Class