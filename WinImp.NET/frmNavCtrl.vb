Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmNavCtrl
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'060304 rjk: Added Sonar parsing to Navigation form.
	'230406 rjk: Fixed the refresh of the units for 4.3.X servers.
	'160706 rjk: Added AutoView for navigate and march.
	
	Public Looking As Boolean
	Public Sonar As Boolean
	
	Public AutoLooking_when_notAutoStopping As Boolean
	Public AutoViewing_when_notAutoStopping As Boolean
	Public RefreshBmapOnStop As Boolean
	Private fooLook As Boolean
	
	Dim TextLines As Short
	Dim InputInhibited As Boolean
	
	Private Sub cmdDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDir.Click
		Dim Index As Short = cmdDir.GetIndex(eventSender)
		'The buttons caption field contains the code to send
		SendCommand(LCase(cmdDir(Index).Text))
		If cmdDir(Index).Text <> "H" Then
			fooLook = True
		Else
			fooLook = False
		End If
		
	End Sub
	
	Public Sub callback()
		'drk@unxsoft.com 7/13/05
		If AutoLooking_when_notAutoStopping And fooLook Then
			Looking = True
			SendCommand("l")
		End If
		If AutoViewing_when_notAutoStopping And fooLook Then
			SendCommand("v")
		End If
		fooLook = False
	End Sub
	
	Private Sub cmdOption_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOption.Click
		Dim Index As Short = cmdOption.GetIndex(eventSender)
		
		'If this is a look, set the looking flag so info will get written to
		'enemy information database
		If cmdOption(Index).Tag = "l" Then
			Looking = True
		End If
		If cmdOption(Index).Tag = "s" Then
			Sonar = True
		End If
		
		'The buttons tag field contains the code to send
		SendCommand(cmdOption(Index).Tag)
	End Sub
	
	Private Sub SendCommand(ByRef strCmd As String)
		
		'Send Command to Server
		frmEmpCmd.SendStringtoServer(strCmd)
		frmEmpCmd.List1.Items.Add(strCmd)
		AddReportLine(strCmd)
		
		EnableButtons(False)
		
		'if this is a movement, set the retreat path
		If InStr("yugjbn", strCmd) > 0 Then
			txtRetreatPath.Text = ReversePath(strCmd) & txtRetreatPath.Text
			If Len(txtRetreatPath.Text) > 5 Then
				txtRetreatPath.Text = VB.Left(txtRetreatPath.Text, 5)
			End If
		End If
		
		'if we are done moving, submit the retreat command
		If VB.Left(strCmd, 1) = "h" And chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked And Len(txtRetreatPath.Text) > 0 And Len(frmDrawMap.LastRetreatUnits) > 0 Then
			strCmd = frmDrawMap.LastRetreatType & "retreat " & frmDrawMap.LastRetreatUnits & " " & txtRetreatPath.Text & " " & VB.Left(Combo1.Text, 1)
			frmEmpCmd.SubmitEmpireCommand(strCmd, False)
		End If
		If VB.Left(strCmd, 1) = "h" Then
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			frmEmpCmd.SubmitEmpireCommand("map *", False)
			frmEmpCmd.SubmitEmpireCommand("bmap *", False)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmNavCtrl.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmNavCtrl_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		'If mimimized for some reason, pop it up
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Normal
			Exit Sub
		End If
		
	End Sub
	Private Sub cmdOption_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdOption.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = cmdOption.GetIndex(eventSender)
		frmNavCtrl_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub cmdDir_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdDir.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = cmdDir.GetIndex(eventSender)
		frmNavCtrl_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmNavCtrl_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'if the form is closed while server is waiting for
		'a response, send an "h" to halt movement
		If frmEmpCmd.ServerQuery Then
			SendCommand("h")
		End If
		
		'drk 7/13/02: is there a better way to do this?
		If RefreshBmapOnStop Then
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			frmEmpCmd.SubmitEmpireCommand("map *", False)
			frmEmpCmd.SubmitEmpireCommand("bmap *", False)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
		
	End Sub
	
	Private Sub lstReport_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lstReport.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		frmNavCtrl_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmNavCtrl_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If InputInhibited Then
			GoTo EventExitSub
		End If
		
		'If a legit command is hit, process it
		If InStr("yughjbnvirslMBfm", Chr(KeyAscii)) > 0 Then
			SendCommand(Chr(KeyAscii))
		End If
		
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmNavCtrl_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
		TextLines = 0
		EnableButtons()
		
		If Len(frmDrawMap.LastRetreatUnits) > 0 Then
			LoadRetreatBox(Combo1)
			Combo1.SelectedIndex = 0
			'If Len(frmDrawMap.LastRetreatString) > 0 Or frmDrawMap.mnuOptions(10).Checked = True Then
			'drk@unxsoft.com 10/25/02: the setting on the frmNavCtrl Box should not be overriden by the options menu.
			'The options menu should provide a default setting for that form.
			If frmDrawMap.DoRetreat Then
				chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkRetreat.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
		Else
			chkRetreat.CheckState = System.Windows.Forms.CheckState.Unchecked
			Frame1.Visible = False
		End If
		
		txtRetreatPath.Text = frmDrawMap.LastRetreatString
	End Sub
	
	'Public Sub ClearReport()   removed efj 8/2003
	''clear text box
	'lstReport.Clear
	'' TextMaxWidth = 0
	'TextLines = 0
	'End Sub
	
	Public Sub AddReportLine(ByRef strLine As String)
		'Add line to text box
		lstReport.Items.Add(strLine)
		TextLines = TextLines + 1
		
	End Sub
	
	
	Public Sub EnableButtons(Optional ByRef bState As Boolean = False)
		
		Dim X As Short
		
		'Default is true
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bState) Then
			bState = True
		End If
		
		'First set direction buttons
		For X = 0 To 6
			cmdDir(X).Enabled = bState
		Next X
		
		'Then set options
		For X = 0 To 9
			cmdOption(X).Enabled = bState
		Next X
		
		InputInhibited = Not bState
		
	End Sub
End Class