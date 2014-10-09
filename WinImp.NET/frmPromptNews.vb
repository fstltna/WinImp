Option Strict Off
Option Explicit On
Friend Class frmPromptNews
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: added Option Explicit and removed dead variables
	'081203 efj: added missing variable definitions
	'092703 rjk: general reformatting
	'113003 rjk: Added Hourly Activity Report.  Added check box for starting relations grid.
	'120203 rjk: Switched output selection to an option box to make clearer.
	
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
		Dim strConnect As String
		
		'make sure number of days are valid
		If Option2.Checked Then
			If Val(cbDays.Text) = 0 Then
				MsgBox("Number of days not valid", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "Entry Error")
				cbDays.Focus()
				Exit Sub
			End If
		End If
		
		'make sure actor name/number
		If chkActor.CheckState = 1 Then
			If Len(cbActor.Text) = 0 Then
				MsgBox("Must choose nation for Actor", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "Entry Error")
				cbVictim.Focus()
				Exit Sub
			End If
		End If
		
		'make sure victim name/number
		If chkVictim.CheckState = 1 Then
			If Len(cbVictim.Text) = 0 Then
				MsgBox("Must choose nation for Victim", MsgBoxStyle.OKOnly + MsgBoxStyle.Exclamation, "Entry Error")
				cbVictim.Focus()
				Exit Sub
			End If
		End If
		
		'112903 rjk: Popup the Relations Grid to show the Telegram results.
		If chkTelegramCount.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			VB6.ShowForm(frmRelationsGrid, VB6.FormShowConstants.Modeless, frmDrawMap)
			frmRelationsGrid.go()
		End If
		
		'112903 rjk: Added By Hour Activity Report
		If optNewsPaper.Checked Then
			frmEmpCmd.SubmitEmpireCommand("np1 Newspaper", False)
		ElseIf optHourly.Checked Then 
			frmEmpCmd.SubmitEmpireCommand("np1 Hourly", False)
		ElseIf optNone.Checked Then 
			frmEmpCmd.SubmitEmpireCommand("np1 None", False)
		End If
		
		'Check if newspaper or just headlines
		If chkHeadlines.CheckState = 1 Then
			strCmd = "headlines"
		Else
			strCmd = "newspaper"
		End If
		
		'Check for number of days
		If Option2.Checked Then
			strCmd = strCmd & " " & CStr(Val(Me.cbDays.Text))
		End If
		
		'Check for optional victim & actor prompts
		If chkActor.CheckState = 1 Then
			strCmd = strCmd & " ?actor=" & CStr(VB6.GetItemData(Me.cbActor, Me.cbActor.SelectedIndex))
			strConnect = "&"
		Else
			strConnect = " ?"
		End If
		
		If chkVictim.CheckState = 1 Then
			strCmd = strCmd & strConnect & "victim=" & CStr(VB6.GetItemData(Me.cbVictim, Me.cbVictim.SelectedIndex))
		End If
		
		'Submit the command
		frmEmpCmd.SubmitEmpireCommand(strCmd)
		
		'112903 rjk: Added By Hour Activity Report
		If Not optNewsPaper.Checked Then
			frmEmpCmd.SubmitEmpireCommand("np2", False)
		End If
		
		Me.Close()
	End Sub
	
	Private Sub frmPromptNews_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim X As Short ' added 8/12/03 efj
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		For X = 1 To 14
			Me.cbDays.Items.Add(CStr(X))
		Next X
		
		cbDays.Text = CStr(1)
		Option1.Checked = True
		
		'Fill list box with nation names
		Nations.FillListBox(cbActor)
		Nations.FillListBox(cbVictim)
	End Sub
	
	Private Sub frmPromptNews_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
End Class