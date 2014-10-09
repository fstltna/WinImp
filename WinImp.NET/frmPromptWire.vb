Option Strict Off
Option Explicit On
Friend Class frmPromptWire
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: added "option explicit"  and definitions for undefined variables
	'081103 efj: removed dead variables
	'093003 rjk: general reformatting.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	
	'UPGRADE_WARNING: Event cbDays.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cbDays.Change was upgraded to cbDays.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cbDays_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbDays.TextChanged
		Option2.Checked = True
	End Sub
	
	'UPGRADE_WARNING: Event cbDays.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbDays_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbDays.SelectedIndexChanged
		Option2.Checked = True
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		' Dim strConnect As String   removed efj 8/2003
		
		'make sure number of days are valid
		If Option2.Checked Then
			If Val(cbDays.Text) = 0 Then
				MsgBox("Number of days not valid", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "Entry Error")
				cbDays.Focus()
				Exit Sub
			End If
		End If
		
		strCmd = "wire"
		
		'Check for number of days
		If Option2.Checked Then
			strCmd = strCmd & " " & CStr(Val(Me.cbDays.Text))
		End If
		
		'Check if clearimg
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			strCmd = strCmd & " y"
		Else
			strCmd = strCmd & " n"
		End If
		
		'Submit the command
		frmEmpCmd.SubmitEmpireCommand(strCmd)
		
		Me.Close()
	End Sub
	
	Private Sub frmPromptWire_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long   removed efj 8/2003
		' success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)      replaced with the following line to remove success 8/2003
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		Dim X As Short ' added definition efj 8/2003
		For X = 1 To 14
			Me.cbDays.Items.Add(CStr(X))
		Next X
		
		cbDays.Text = CStr(1)
		Option1.Checked = True
		
		'Set Check box
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	Private Sub frmPromptWire_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
End Class