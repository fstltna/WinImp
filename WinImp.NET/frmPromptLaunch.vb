Option Strict Off
Option Explicit On
Friend Class frmPromptLaunch
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092803 rjk: Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
	'            General reformatting.
	'180206 rjk: Replace pdump and lost with GetPlanes and GetLost.
	
	Public strCmd As String
	Public ShowRange As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		strCmd = strCmd & " " & txtUnit.Text
		frmEmpCmd.SubmitEmpireCommand("la1 " & txtPath.Text, False)
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetPlanes()
		GetLost()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptLaunch_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		'rjk 8/11/03 Added LastLoad show the grid gets updated with inrange missiles each time the Launch form is launched
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptLaunch_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		If Len(txtUnit.Text) = 0 Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Len(txtPath.Text) > 0 And ShowRange Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_P_MISSILE, True, True, True, txtPath.Text)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PC_LAUNCH, True, True, False)
		End If
	End Sub
End Class