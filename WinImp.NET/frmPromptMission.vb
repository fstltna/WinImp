Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptMission
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'120803 efj: removed dead variables
	'270903 rjk: general reformatting
	'030206 rjk: request an update of the mission values when changing a mission.
	'180206 rjk: Replace mission with xdump for 4.3.0 servers.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strType As String
		
		If Option1.Checked Then
			strType = "ship "
		ElseIf Option2.Checked Then 
			If Len(Trim(txtRange.Text)) > 0 Then
				strCmd = "lrange " & txtUnit.Text & " " & txtRange.Text
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
			strType = "land "
		Else
			strType = "plane "
		End If
		
		strCmd = "mission " & strType & txtUnit.Text & " " & VB.Left(Combo1.Text, 1) & " " & txtOrigin.Text & " " & txtNum.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			strCmd = "mission " & strType & txtUnit.Text & " q"
		Else
			strCmd = "xdump " & strType & txtUnit.Text
		End If
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub frmPromptMission_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptMission_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Combo1.Items.Clear()
			Combo1.Items.Add("interdiction")
			Combo1.Items.Add("clear")
			Combo1.Items.Add("query")
			Combo1.SelectedIndex = 0
			Label6.Visible = False
			txtRange.Visible = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			Combo1.Items.Clear()
			Combo1.Items.Add("interdiction")
			Combo1.Items.Add("reserve")
			Combo1.Items.Add("clear")
			Combo1.Items.Add("query")
			Combo1.SelectedIndex = 0
			Label6.Visible = True
			txtRange.Visible = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			Combo1.Items.Clear()
			Combo1.Items.Add("interdiction")
			Combo1.Items.Add("support")
			Combo1.Items.Add("off support")
			Combo1.Items.Add("def support")
			Combo1.Items.Add("escort")
			Combo1.Items.Add("air defense")
			Combo1.Items.Add("clear")
			Combo1.Items.Add("query")
			Combo1.SelectedIndex = 0
			Label6.Visible = False
			txtRange.Visible = False
		End If
	End Sub
End Class