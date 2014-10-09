Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmScript
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead procedures and variables
	'093003 rjk: general reformatting.
	'110103 rjk: Form cleanup - removed left and width from toolbar.
	
	Dim Dirty As Boolean
	
	'Private Sub cmdCancel_Click()    8/2003 efj  removed
	'If CheckDirty Then
	'    Unload Me
	'End If
	'End Sub
	
	Private Sub ButtonRecord()
		'if currently recording - process as a stop
		If frmEmpCmd.RecordingScriptFile Then
			frmEmpCmd.RecordingScriptFile = False
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).Text = "Rec."
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).ImageIndex = 5
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).ToolTipText = "Start Script Recorder"
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(8).Enabled = True
		Else
			'start recording
			frmEmpCmd.RecordingScriptFile = True
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).Text = "Stop"
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).ImageIndex = 6
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(6).ToolTipText = "Stop Script Recorder"
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(8).Enabled = False
		End If
	End Sub
	
	Private Sub frmScript_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'UPGRADE_ISSUE: MSComctlLib.Toolbar property tbToolBar.ButtonWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.tbToolBar.Left) + tbToolBar.ButtonWidth * 8)
		Me.rt1.Width = Me.ClientRectangle.Width
	End Sub
	
	'UPGRADE_WARNING: Event frmScript.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmScript_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Exit Sub
		End If
		
		Me.rt1.Width = Me.ClientRectangle.Width
		Me.rt1.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(Me.rt1.Top))
	End Sub
	
	Private Sub frmScript_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'Make sure scripting is turned off
		frmEmpCmd.RecordingScriptFile = False
	End Sub
	
	Private Function CheckDirty() As Boolean
		Dim Reply As Short
		'if file has been changed since last save, update
		If Dirty Then
			Reply = MsgBox("File has changed.  Do you want to save changes ?", MsgBoxStyle.YesNoCancel, "Confirm Clear")
			If Reply = MsgBoxResult.Cancel Then
				CheckDirty = False
				Exit Function
			End If
			
			If Reply = MsgBoxResult.No Then
				CheckDirty = True
				Exit Function
			End If
			
			'save changes
			CheckDirty = ButtonSave
		Else
			CheckDirty = True
		End If
	End Function
	
	Private Sub rt1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles rt1.TextChanged
		Dirty = True
	End Sub
	
	Private Sub tbToolBar_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _tbToolBar_Button1.Click, _tbToolBar_Button2.Click, _tbToolBar_Button3.Click, _tbToolBar_Button4.Click, _tbToolBar_Button5.Click, _tbToolBar_Button6.Click, _tbToolBar_Button7.Click, _tbToolBar_Button8.Click, _tbToolBar_Button9.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		' Dim success As Long
		'turn off topmost
		Call SetWindowPos(Me.Handle.ToInt32, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
		Select Case Button.Name
			Case "New"
				ButtonNew()
			Case "Open"
				ButtonOpen()
			Case "Save"
				ButtonSave()
			Case "Exit"
				Me.Close()
			Case "Record"
				ButtonRecord()
			Case "Run"
				ButtonRun()
			Case "Timer"
				ButtonTimer()
		End Select
		
		'Pop topmost back up
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
	End Sub
	
	Private Function ButtonSave() As Boolean
		Dim strFile As String
		
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		cd1.CancelError = True
		On Error GoTo ErrHandler
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property cd1.Flags was upgraded to cd1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		cd1Open.ShowReadOnly = False
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		cd1Open.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		cd1Save.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		' Specify default filter
		cd1Open.FilterIndex = 2
		cd1Save.FilterIndex = 2
		' Set default file name
		'cd1.FileName = "Empire Script.txt" 'drk@unxsoft.com.  this annoyed me. did it annoy anyone else, or should I put it back?
		
		If VerifySubDirectory("Scripts", True) Then
			cd1Open.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
			cd1Save.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
		End If
		
		' Display the Open dialog box
		cd1Open.Title = "Save Script File"
		cd1Save.Title = "Save Script File"
		cd1Save.ShowDialog()
		cd1Open.FileName = cd1Save.FileName
		' Display name of selected file
		
		On Error GoTo 0
		
		Dim Reply As Short
		
		strFile = cd1Open.FileName
		
		'Append file suffix if necessary
		If InStr(strFile, ".") = 0 Then
			strFile = strFile & ".txt"
		End If
		
		'If already exists, then prompt
		Reply = MsgBoxResult.No
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(strFile)) > 0 Then
			
			Reply = MsgBox("File " & strFile & " alreay exists.  Do you wish to " & "replace it?", MsgBoxStyle.OKCancel + MsgBoxStyle.Question, "File Exists")
			
			If Reply = MsgBoxResult.Cancel Then
				ButtonSave = False
				Exit Function
			End If
		End If
		
		rt1.SaveFile(strFile, Windows.Forms.RichTextBoxStreamType.PlainText)
		Dirty = False
		ButtonSave = True
		Exit Function
		
ErrHandler: 
		'User pressed the Cancel button
		ButtonSave = False
	End Function
	
	
	Private Sub ButtonOpen()
		Dim strFile As String
		
		If Not CheckDirty Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		cd1.CancelError = True
		On Error GoTo ErrHandler
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property cd1.Flags was upgraded to cd1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		cd1Open.ShowReadOnly = False
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		cd1Open.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		cd1Save.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		' Specify default filter
		cd1Open.FilterIndex = 2
		cd1Save.FilterIndex = 2
		
		' Display the Open dialog box
		cd1Open.Title = "Open Script File"
		cd1Save.Title = "Open Script File"
		cd1Open.FileName = vbNullString
		cd1Save.FileName = vbNullString
		
		If VerifySubDirectory("Scripts", True) Then
			cd1Open.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
			cd1Save.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
		End If
		
		cd1Open.ShowDialog()
		cd1Save.FileName = cd1Open.FileName
		
		On Error GoTo 0
		
		' Dim strReply As Integer    8/2003 efj  removed
		
		strFile = cd1Open.FileName
		
		rt1.LoadFile(strFile, Windows.Forms.RichTextBoxStreamType.PlainText)
		Exit Sub
		
ErrHandler: 
		'User pressed the Cancel button
	End Sub
	Private Sub ButtonNew()
		If Not CheckDirty Then
			Exit Sub
		End If
		
		'Clear the text file
		rt1.Text = vbNullString
		Dirty = False
	End Sub
	
	Private Sub ButtonRun()
		Dim strSource As String
		Dim strTemp As String
		Dim n As Short
		
		frmEmpCmd.SubmitEmpireCommand("bf1", False)
		strSource = rt1.Text
		Do 
			n = InStr(strSource, vbLf)
			If n = 0 Then n = Len(strSource) + 1
			strTemp = VB.Left(strSource, n - 1)
			strSource = Mid(strSource, n + 1)
			
			'remove carrage returns
			n = InStr(strTemp, vbCr)
			While n > 0
				strTemp = VB.Left(strTemp, n - 1) & Mid(strTemp, n + 1)
				n = InStr(strTemp, vbCr)
			End While
			
			'Execute the command
			frmEmpCmd.SubmitEmpireCommand(strTemp, False)
			
		Loop While Len(strSource) > 0
		
		frmEmpCmd.SubmitEmpireCommand("bf2", False)
	End Sub
	
	Private Sub ButtonTimer()
		Dim strFile As String
		Dim nReply As Short
		Dim strReply As String
		
		strReply = TimerHeader
		strReply = strReply & vbCr & vbCr & "Do you wish to change selections? " & vbCr & vbCr & "Select Yes to Change" & vbCr & "Select No to Exit without Changes" & vbCr & "Select Cancel to Delete Current Selections"
		
		nReply = MsgBox(strReply, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "Current Timed Script Settings")
		
		'If cancel, wipe out scripts
		If nReply = MsgBoxResult.Cancel Then
			If Len(BatchScript1File) > 0 Or Len(BatchScriptUpdate) > 0 Then
				BatchScript1File = vbNullString
				BatchScriptUpdate = vbNullString
				BatchScript1Time = 0
				strReply = TimerHeader & vbCr & vbCr & "Scipts cleared"
				MsgBox(strReply, MsgBoxStyle.OKOnly, "Timed Script Settings")
			End If
		End If
		
		If nReply <> MsgBoxResult.Yes Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
		cd1.CancelError = True
		On Error GoTo ButtonTimerExit
		' Set flags
		'UPGRADE_WARNING: MSComDlg.CommonDialog property cd1.Flags was upgraded to cd1Open.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		'UPGRADE_WARNING: FileOpenConstants constant FileOpenConstants.cdlOFNHideReadOnly was upgraded to OpenFileDialog.ShowReadOnly which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		cd1Open.ShowReadOnly = False
		' Set filters
		'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		cd1Open.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		cd1Save.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
		' Specify default filter
		cd1Open.FilterIndex = 2
		cd1Save.FilterIndex = 2
		
		' Display the Open dialog box
		cd1Open.Title = "Select Script File"
		cd1Save.Title = "Select Script File"
		cd1Open.FileName = vbNullString
		cd1Save.FileName = vbNullString
		If VerifySubDirectory("Scripts", True) Then
			cd1Open.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
			cd1Save.InitialDirectory = My.Application.Info.DirectoryPath & "\Scripts"
		End If
		
		cd1Open.ShowDialog()
		cd1Save.FileName = cd1Open.FileName
		
		On Error GoTo 0
		
		strFile = cd1Open.FileName
		
		nReply = MsgBox("Run script " & strFile & " right after update?", MsgBoxStyle.YesNoCancel, "Script Timer")
		If nReply = MsgBoxResult.Cancel Then
			GoTo ButtonTimerExit
		End If
		
		If nReply = MsgBoxResult.Yes Then
			BatchScriptUpdate = strFile
			GoTo ButtonTimerExit
		End If
		
		Do 
			strReply = InputBox("Run time prior to update (in seconds)")
			If Val(strReply) <= 0 Then
				nReply = MsgBox("Input invalid - time script cancelled", MsgBoxStyle.RetryCancel + MsgBoxStyle.Exclamation)
				If nReply = MsgBoxResult.Cancel Then
					Exit Sub
				Else
					strReply = vbNullString
				End If
			End If
		Loop Until Val(strReply) > 0
		
		BatchScript1File = strFile
		BatchScript1Time = Val(strReply)
		
		
		'User pressed the Cancel button
ButtonTimerExit: 
		
		'display final settings
		strReply = "Current Settings:" & vbCr & vbCr & TimerHeader & vbCr
		MsgBox(strReply, MsgBoxStyle.OKOnly, "Timed Script Settings")
	End Sub
	
	Private Function TimerHeader() As String
		Dim strTemp As String
		strTemp = "Script to run "
		If Len(BatchScript1File) > 0 Then
			strTemp = strTemp & "at update minus " & CStr(BatchScript1Time) & " seconds: " & vbCr & BatchScript1File
		Else
			strTemp = strTemp & "prior to update: None"
		End If
		strTemp = strTemp & vbCr & vbCr & "Script to run post update: "
		If Len(BatchScriptUpdate) > 0 Then
			strTemp = strTemp & vbCr & BatchScriptUpdate
		Else
			strTemp = strTemp & "None"
		End If
		
		TimerHeader = strTemp
	End Function
	
	Public Sub AddCommand(ByRef strCmd As String)
		rt1.Text = rt1.Text & strCmd & vbLf
	End Sub
	
	Public Sub ExectuteBatchScript(ByRef strFile As String)
		' Dim FileName As String    8/2003 efj  removed
		
		rt1.LoadFile(strFile, Windows.Forms.RichTextBoxStreamType.PlainText)
		ButtonRun()
		Me.Close()
	End Sub
End Class