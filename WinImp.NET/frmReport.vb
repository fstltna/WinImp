Option Strict Off
Option Explicit On
Friend Class frmReport
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'300903 rjk: General reformatting.
	'290804 rjk: Always save telegrams unless it is duplicate.
	'170906 rjk: Added TTS report window.
	'020907 rjk: Switched to database field for TimeZoneAdj.
	
	Private TextMaxWidth As Short
	Private TextLines As Short
	Private SavedOnce As Boolean
	
	'UPGRADE_WARNING: Form event frmReport.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmReport_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim nLen As Integer
		
		'If mimimized for some reason, pop it up
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Normal
			Exit Sub
		End If
		
		If TextMaxWidth < 5340 Then
			TextMaxWidth = 5340
		End If
		
		If TextMaxWidth > 0 Then
			If TextMaxWidth > ((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 4) / 5) Then
				TextMaxWidth = (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 4) / 5
			End If
			
			If Me.WindowState <> System.Windows.Forms.FormWindowState.Maximized Then
				Me.Width = VB6.TwipsToPixelsX(TextMaxWidth * 1.15)
				Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
			End If
		End If
		
		If TextLines > 1 Then
			'UPGRADE_ISSUE: Form method frmReport.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			nLen = Me.TextHeight("XX") * (TextLines + 2)
			
			
			Text1.Height = VB6.TwipsToPixelsY(nLen)
			If VB6.PixelsToTwipsY(Text1.Height) > ((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4) Then
				Text1.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4)
			End If
			
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Text1.Height) + VB6.PixelsToTwipsY(tbToolBar.Height) * 2)
			Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
		End If
	End Sub
	
	Private Sub frmReport_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
		TextLines = 0
		SavedOnce = False
	End Sub
	
	'UPGRADE_WARNING: Event frmReport.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmReport_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If Me.WindowState <> System.Windows.Forms.FormWindowState.Normal Then
			Exit Sub
		End If
		
		'Check for minimum width, height
		If VB6.PixelsToTwipsX(Me.Width) < 4005 Then
			Me.Width = VB6.TwipsToPixelsX(4005)
		End If
		
		If VB6.PixelsToTwipsY(Me.Height) < 6675 Then
			Me.Height = VB6.TwipsToPixelsY(6675)
		End If
		
		'Position tool bar
		tbToolBar.SetBounds(0, 0, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		'Position text Box
		'lstReport.Move 0, tbToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - tbToolBar.Height
		Text1.SetBounds(0, tbToolBar.Height, Me.ClientRectangle.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - VB6.PixelsToTwipsY(tbToolBar.Height)))
	End Sub
	
	Public Sub ClearReport()
		'clear text box
		'lstReport.Clear
		Text1.Text = vbNullString
		TextMaxWidth = 0
		TextLines = 0
	End Sub
	
	Public Sub AddReportLine(ByRef strLine As String)
		' Dim S() As String    8/2003 efj  removed
		Dim X, nLen As Short
		
		On Error GoTo AddReportLine_exit
		
		If frmEmpCmd.SubmittedFromCommandLine Then
			Exit Sub
		End If
		
		
		'Add line to text box
		'lstReport.AddItem strLine
		Text1.Text = Text1.Text & vbCrLf & strLine
		TextLines = TextLines + 1
		
		'UPGRADE_ISSUE: Form method frmReport.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		X = frmReport.TextWidth(strLine)
		
		If X > TextMaxWidth Then
			TextMaxWidth = X
		End If
		
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		End If
		
		If TextLines > 1 Then
			'UPGRADE_ISSUE: Form method frmReport.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			nLen = Me.TextHeight("XX") * (TextLines + 2)
			'If lstReport.Height < nLen Then
			If VB6.PixelsToTwipsY(Text1.Height) < nLen Then
				
				'lstReport.Height = nLen
				Text1.Height = VB6.TwipsToPixelsY(nLen)
				'If lstReport.Height > ((Screen.Height * 3) / 4) Then
				'    lstReport.Height = (Screen.Height * 3) / 4
				'End If
				If VB6.PixelsToTwipsY(Text1.Height) > ((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4) Then
					Text1.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4)
				End If
				
				'Me.Height = lstReport.Height + tbToolBar.Height * 2
				Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Text1.Height) + VB6.PixelsToTwipsY(tbToolBar.Height) * 2)
				Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
			End If
		End If
		
		If frmOptions.bTTSReportWindow And Len(strLine) > 0 Then
			'    ttsReport.Tag = "Normal"
			'    ttsReport.Speak strLine
		End If
		
AddReportLine_exit: 
	End Sub
	
	'Private Sub lstReport_KeyPress(KeyAscii As Integer)    8/2003 efj  removed
	'
	'Dim txtstring As String
	'
	''Trap Control C and Copy to Clipboard
	'If KeyAscii = 3 Then
	'    Clipboard.SetText Text1.Text
	'End If
	'
	'End Sub
	
	Private Sub tbToolBar_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _tbToolBar_Button1.Click, _tbToolBar_Button2.Click, _tbToolBar_Button3.Click, _tbToolBar_Button4.Click, _tbToolBar_Button5.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		Select Case Button.Name
			Case "Save"
				ButtonSave()
			Case "Copy"
				My.Computer.Clipboard.SetText(Me.Text1)
			Case "File"
				ButtonFile()
			Case "Exit"
				Me.Close()
		End Select
	End Sub
	
	Private Sub ButtonFile()
		If SavedOnce Then
			Exit Sub
		End If
		
		' Copy Selected text into a sring
		IncomingTelegramLine("> Saved " & Me.Text & "  dated " & VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, rsVersion.Fields("Time Zone Adj").Value, Now), "ddd mmm dd hh:mm:ss yyyy"), True)
		' Dim x As Integer    8/2003 efj  removed
		'For x = 0 To (lstReport.ListCount - 1)
		'    IncomingTelegramLine lstReport.List(x), False, True
		'Next x
		IncomingTelegramLine(Text1.Text, False)
		SavedOnce = True
	End Sub
	
	Private Sub ButtonSave()
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
		cd1Open.FileName = Me.Text
		cd1Save.FileName = Me.Text
		
		If VerifySubDirectory("Saved Reports", True) Then
			cd1Open.InitialDirectory = My.Application.Info.DirectoryPath & "\Saved Reports"
			cd1Save.InitialDirectory = My.Application.Info.DirectoryPath & "\Saved Reports"
		End If
		
		' Display the Open dialog box
		cd1Save.ShowDialog()
		cd1Open.FileName = cd1Save.FileName
		' Display name of selected file
		
		On Error GoTo 0
		
		Dim strFile As String 'drk@unxsoft.com 7/03/02
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
			
			Reply = MsgBox("File " & strFile & " alreay exists.  Do you wish to " & "append this report to the file ?  (Take Yes to Append, No to Replace, " & "Cancel to Exit", MsgBoxStyle.YesNoCancel, "File Exists")
			
			If Reply = MsgBoxResult.Cancel Then
				Exit Sub
			End If
		End If
		
		Dim filenum As Short
		' Dim x As Integer    8/2003 efj  removed
		
		filenum = FreeFile
		If Reply = MsgBoxResult.Yes Then
			FileOpen(filenum, strFile, OpenMode.Append)
		Else
			FileOpen(filenum, strFile, OpenMode.Output)
		End If
		
		PrintLine(filenum, Text1.Text)
		
		FileClose(filenum)
		Exit Sub
		
ErrHandler: 
		'User pressed the Cancel button
		Exit Sub
	End Sub
End Class