Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmTelegram
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead variables
	'093003 rjk: general reformatting and cleanup
	'101503 rjk: Fixed so telegram does not delete blank lines.
	'103003 rjk: Added Import Intelligence functionality
	'110303 rjk: Modified ButtonSend to account for the server not supporting up 1024 characters,
	'            it only supports 1022 characters in a telegram line.
	'            Fixed to support lines longer than 2044.
	'111403 rjk: Clear Telegram window after removing all the telegrams.
	'111803 rjk: Switched to a Yes/No question box for clearing telegrams.
	'112903 rjk: Remove the question on clearing the telegrams as there is a separate form now.
	'270303 rjk: Switched over to DeleteAllRecords for clearing tables.
	'110804 rjk: Added the motd to telegram window for deity mode.
	'290804 rjk: Always save telegrams, unless it is a duplicate.
	'250505 rjk: Added Telegram Warning, Soft Limit and Hard Limit.
	'180206 rjk: Removed "The deity" from the telegram list for 4.3.0.  The 4.3.0 country list now
	'            contains the deities in it.
	'181006 rjk: Added Import Offset to ImportIntelligence.
	'090607 rjk: Added forward ability for telegrams.
	'010108 rjk: Fixed incorrect telegram destinations.
	
	Const NAME_COLUMN As Short = 0
	Const TYPE_COLUMN As Short = 1
	Const SIZE_COLUMN As Short = 2
	Const DATE_COLUMN As Short = 3
	
	Const BUTTON_FIRST As Short = 1
	Const BUTTON_BACK As Short = 2
	Const BUTTON_NEXT As Short = 3
	Const BUTTON_LATEST As Short = 4
	'Separator
	Const BUTTON_NEW As Short = 6
	Const BUTTON_REPLY As Short = 7
	Const BUTTON_FORWARD As Short = 8
	Const BUTTON_IMPORT As Short = 9 '103003 rjk: Added Import Intelligence functionality
	'Separator
	Const BUTTON_DELETE As Short = 11
	Const BUTTON_SEND As Short = 12
	
	Public TreeDirty As Boolean
	
	Dim NewTelegram As Boolean
	Dim SelectedTelegram As Short
	' Dim hldNation As String    8/2003 efj  removed
	Dim CurTele As Short
	
	Dim mbMoving As Boolean
	Const sglSplitLimit As Short = 500
	
	Enum eMessageCreate
		MESS_NEW
		MESS_REPLY
		MESS_FORWARD
	End Enum
	Dim bForwardMessage As Boolean
	
	'UPGRADE_WARNING: Event chkInclude.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkInclude_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkInclude.CheckStateChanged
		Dim hldtxt As String
		Dim strBuffer As String
		Dim n As Short
		Dim strPlayerName As String
		
		If chkInclude.CheckState = System.Windows.Forms.CheckState.Checked Then
			'copy text of letter into the top of telegram
			rsTeleBody.Seek("=", SelectedTelegram, 1)
			
			'Load telegram body into text box
			hldtxt = txtBody.Text
			txtBody.Text = vbNullString
			If Not rsTeleBody.NoMatch Then
				If bForwardMessage Then
					txtBody.Text = txtBody.Text & "}" & vbLf
					strPlayerName = Nations.NationName(rsTeleHead.Fields("From").Value)
					If Len(strPlayerName) = 0 Then
						strPlayerName = "Nation #" & CStr(rsTeleHead.Fields("From").Value)
					End If
					txtBody.Text = txtBody.Text & "}From: " & strPlayerName & vbLf & "}" & vbLf
				End If
				Do While (Not rsTeleBody.EOF)
					If rsTeleBody.Fields("ID").Value <> SelectedTelegram Then
						Exit Do
					End If
					txtBody.Text = txtBody.Text & "}" + rsTeleBody.Fields("Body").Value + vbLf
					rsTeleBody.MoveNext()
				Loop 
			End If
			txtBody.Text = txtBody.Text & vbLf & hldtxt
		Else
			'remove all the copied lines in the text
			strBuffer = txtBody.Text
			txtBody.Text = vbNullString
			While Len(strBuffer) > 0
				'Buffer could contain several lines, so seperate at line feeds
				n = InStr(strBuffer, vbLf)
				If n = 0 Then
					hldtxt = strBuffer
					strBuffer = vbNullString
				Else
					hldtxt = VB.Left(strBuffer, n - 1)
					If Len(strBuffer) = n Then
						strBuffer = vbNullString
					Else
						strBuffer = Mid(strBuffer, n + 1)
					End If
				End If
				If VB.Left(hldtxt, 1) <> "}" Then
					txtBody.Text = txtBody.Text & hldtxt & vbLf
				End If
			End While
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmTelegram.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmTelegram_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
		If TreeDirty And Not NewTelegram Then
			LoadTree()
			'Set selection to last item viewed
			If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
				rsTeleHead.MoveLast()
				CurTele = rsTeleHead.Fields("ID").Value
				tvTreeView.SelectedNode = tvTreeView.Nodes.Item("T" & CStr(CurTele))
				FillTextBox(tvTreeView.SelectedNode)
				tvTreeView.SelectedNode.EnsureVisible()
			End If
		End If
		If bDeity Then
			optMotd.Visible = True
		End If
	End Sub
	
	Private Sub frmTelegram_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'check and see if the control key is down
		If Not ((Shift And VB6.ShiftConstants.CtrlMask) = 2) Then
			Exit Sub
		End If
		On Error Resume Next
		
		If KeyCode = System.Windows.Forms.Keys.D1 Then
			frmEmpCmd.WindowState = System.Windows.Forms.FormWindowState.Normal
			frmEmpCmd.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D2 Then
			Me.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D3 Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Normal
			Me.Activate()
		End If
	End Sub
	
	Private Sub frmTelegram_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.75)
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.75)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
		imgSplitter.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) \ 3)
		lbNation.Visible = False
		chkInclude.Visible = False
		TreeDirty = True
	End Sub
	
	' 091903 rjk: Removed as there is no code inside
	'Private Sub Form_Paint()
	'txtBody.View = Val(GetSetting(APPNAME, "Settings", "ViewMode", "0"))
	'tbToolBar.Buttons(txtBody.View + LISTVIEW_BUTTON).Value = tbrPressed
	'mnuListViewMode(txtBody.View).Checked = True
	'End Sub
	
	Private Sub frmTelegram_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		'Only unload if form Draw says OK
		'UPGRADE_ISSUE: Constant vbFormCode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Or UnloadMode = vbFormCode Then
			If Not frmDrawMap.ShutDown Then
				Cancel = 1
				Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
			End If
		End If
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub frmTelegram_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Dim i As Short
		
		'close all sub forms
		For i = My.Application.OpenForms.Count - 1 To 1 Step -1
			'UPGRADE_ISSUE: Unload Forms() was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
			Unload(My.Application.OpenForms(i))
		Next 
	End Sub
	
	
	'Private Sub mnuViewOptions_Click()
	'    frmOptions.Show vbModal, Me
	'End Sub
	
	'UPGRADE_WARNING: Event frmTelegram.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmTelegram_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		'If Me.WindowState <> vbNormal Then
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Exit Sub
		End If
		
		On Error Resume Next
		If VB6.PixelsToTwipsX(Me.Width) < 3000 Then Me.Width = VB6.TwipsToPixelsX(3000)
		SizeControls((VB6.PixelsToTwipsX(imgSplitter.Left)))
	End Sub
	
	Private Sub imgSplitter_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgSplitter.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		With imgSplitter
			picSplitter.SetBounds(.Left, .Top, VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) \ 2), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(.Height) - 20))
		End With
		picSplitter.Visible = True
		mbMoving = True
	End Sub
	
	Private Sub imgSplitter_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgSplitter.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim sglPos As Single
		
		If mbMoving Then
			sglPos = X + VB6.PixelsToTwipsX(imgSplitter.Left)
			If sglPos < sglSplitLimit Then
				picSplitter.Left = VB6.TwipsToPixelsX(sglSplitLimit)
			ElseIf sglPos > VB6.PixelsToTwipsX(Me.Width) - sglSplitLimit Then 
				picSplitter.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - sglSplitLimit)
			Else
				picSplitter.Left = VB6.TwipsToPixelsX(sglPos)
			End If
		End If
	End Sub
	
	Private Sub imgSplitter_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles imgSplitter.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		SizeControls((VB6.PixelsToTwipsX(picSplitter.Left)))
		picSplitter.Visible = False
		mbMoving = False
	End Sub
	
	Private Sub SizeControls(ByRef X As Single)
		On Error Resume Next
		
		'set the width
		If X < 1500 Then X = 1500
		If X > (VB6.PixelsToTwipsX(Me.Width) - 1500) Then X = VB6.PixelsToTwipsX(Me.Width) - 1500
		tvTreeView.Width = VB6.TwipsToPixelsX(X)
		imgSplitter.Left = VB6.TwipsToPixelsX(X)
		txtBody.Left = VB6.TwipsToPixelsX(X + 40)
		txtBody.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - (VB6.PixelsToTwipsX(tvTreeView.Width) + 140))
		lblTitle(0).Width = tvTreeView.Width
		lblTitle(1).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtBody.Left) + 20)
		lblTitle(1).Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(txtBody.Width) - 40)
		
		'set the top
		If tbToolBar.Visible Then
			tvTreeView.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(tbToolBar.Height) + VB6.PixelsToTwipsY(picTitles.Height))
		Else
			tvTreeView.Top = picTitles.Height
		End If
		txtBody.Top = tvTreeView.Top
		
		'set the height
		If sbStatusBar.Visible Then
			tvTreeView.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - (VB6.PixelsToTwipsY(picTitles.Top) + VB6.PixelsToTwipsY(picTitles.Height) + VB6.PixelsToTwipsY(sbStatusBar.Height)))
		Else
			tvTreeView.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - (VB6.PixelsToTwipsY(picTitles.Top) + VB6.PixelsToTwipsY(picTitles.Height)))
		End If
		
		'Move the hidden forms
		lbNation.SetBounds(tvTreeView.Left, tvTreeView.Top, tvTreeView.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		labNation.SetBounds(tvTreeView.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(tvTreeView.Top) + VB6.PixelsToTwipsY(lbNation.Height) / 2), tvTreeView.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		chkInclude.SetBounds(tvTreeView.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lbNation.Top) + VB6.PixelsToTwipsY(lbNation.Height)), tvTreeView.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		chkSave.SetBounds(tvTreeView.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(chkInclude.Top) + VB6.PixelsToTwipsY(chkInclude.Height)), tvTreeView.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		frameSend.SetBounds(tvTreeView.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(chkSave.Top) + VB6.PixelsToTwipsY(chkSave.Height)), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		txtBody.Height = tvTreeView.Height
		imgSplitter.Top = tvTreeView.Top
		imgSplitter.Height = tvTreeView.Height
	End Sub
	
	'UPGRADE_WARNING: Event lbNation.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lbNation_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lbNation.SelectedIndexChanged
		SetRecipientLabel()
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			lbNation.Visible = True
			labNation.Visible = False
			'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
			Timer1.Interval = 0
			'Set titles
			lblTitle(0).Text = "Send Telegram To:"
			SetRecipientLabel()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			lbNation.Visible = False
			labNation.Visible = True
			Timer1.Interval = 1000
			'Set titles
			lblTitle(0).Text = "Make Announcement:"
			lblTitle(1).Text = "Enter New Anno"
			labNation.Text = "<< ANNOUNCEMENT >>"
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			lbNation.Visible = True
			labNation.Visible = False
			'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
			Timer1.Interval = 0
			'Set titles
			lblTitle(0).Text = "Send Flash To:"
			SetRecipientLabel()
		End If
	End Sub
	
	
	'UPGRADE_WARNING: Event optMotd.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optMotd_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMotd.CheckedChanged
		If eventSender.Checked Then
			lbNation.Visible = False
			labNation.Visible = True
			Timer1.Interval = 1000
			'Set titles
			lblTitle(0).Text = ""
			lblTitle(1).Text = "Enter New motd"
			labNation.Text = "<< motd >>"
		End If
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		If Option2.Checked Or optMotd.Checked Then
			If labNation.Visible Then
				labNation.Visible = False
				Timer1.Interval = 500
			Else
				labNation.Visible = True
				Timer1.Interval = 1000
			End If
		End If
	End Sub
	
	'UPGRADE_ISSUE: MSComctlLib.TreeView event tvTreeView.DragDrop was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub tvTreeView_DragDrop(ByRef Source As System.Windows.Forms.Control, ByRef X As Single, ByRef Y As Single)
		If Source.equals(imgSplitter.Image) Then
			SizeControls(X)
		End If
	End Sub
	
	Private Sub tbToolBar_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _tbToolBar_Button1.Click, _tbToolBar_Button2.Click, _tbToolBar_Button3.Click, _tbToolBar_Button4.Click, _tbToolBar_Button5.Click, _tbToolBar_Button6.Click, _tbToolBar_Button7.Click, _tbToolBar_Button8.Click, _tbToolBar_Button9.Click, _tbToolBar_Button10.Click, _tbToolBar_Button11.Click, _tbToolBar_Button12.Click, _tbToolBar_Button13.Click, _tbToolBar_Button14.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		' Dim strKey As String    8/2003 efj  removed
		' Dim teleID As Integer    8/2003 efj  removed
		
		rsTeleHead.Index = "PrimaryKey"
		
		Select Case Button.Name
			Case "First"
				ButtonFirst()
			Case "Last"
				ButtonLast()
			Case "Previous"
				ButtonBack()
			Case "Next"
				ButtonNext()
			Case "Delete"
				If NewTelegram Then
					ButtonCancel()
				Else
					ButtonDelete()
				End If
			Case "New"
				ButtonNew(eMessageCreate.MESS_NEW)
			Case "Reply"
				ButtonNew(eMessageCreate.MESS_REPLY)
			Case "Forward"
				ButtonNew(eMessageCreate.MESS_FORWARD)
			Case "Send"
				ButtonSend()
			Case "Import" '103003 rjk: Added Import Intelligence
				ButtonImport()
		End Select
	End Sub
	
	Private Sub ButtonBack()
		Dim strKey As String
		Dim teleID As Short
		
		On Error GoTo TelegramError
		
		strKey = tvTreeView.SelectedNode.Name
		If VB.Left(strKey, 1) <> "T" Then
			Exit Sub
		End If
		teleID = CShort(Mid(strKey, 2))
		rsTeleHead.Seek("=", teleID)
		If Not rsTeleHead.NoMatch Then
			rsTeleHead.MovePrevious()
		End If
		If rsTeleHead.NoMatch Or rsTeleHead.BOF Then
			MsgBox("No More Telegrams", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Back Error")
		Else
			teleID = rsTeleHead.Fields("ID").Value
			tvTreeView.SelectedNode = tvTreeView.Nodes.Item("T" & CStr(teleID))
			FillTextBox(tvTreeView.SelectedNode)
		End If
		
		Exit Sub
TelegramError: 
		TelegramError()
	End Sub
	
	Private Sub ButtonNext()
		Dim strKey As String
		Dim teleID As Short
		
		On Error GoTo TelegramError
		
		strKey = tvTreeView.SelectedNode.Name
		If VB.Left(strKey, 1) <> "T" Then
			Exit Sub
		End If
		teleID = CShort(Mid(strKey, 2))
		rsTeleHead.Seek("=", teleID)
		If Not rsTeleHead.NoMatch Then
			rsTeleHead.MoveNext()
		End If
		If rsTeleHead.NoMatch Or rsTeleHead.EOF Then
			MsgBox("No More Telegrams", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Forward Error")
		Else
			teleID = rsTeleHead.Fields("ID").Value
			tvTreeView.SelectedNode = tvTreeView.Nodes.Item("T" & CStr(teleID))
			FillTextBox(tvTreeView.SelectedNode)
		End If
		
		Exit Sub
TelegramError: 
		TelegramError()
	End Sub
	
	Private Sub ButtonFirst()
		' Dim strKey As String    8/2003 efj  removed
		Dim teleID As Short
		
		On Error GoTo TelegramError
		
		If rsTeleHead.BOF And rsTeleHead.EOF Then
			Exit Sub
		End If
		rsTeleHead.MoveFirst()
		teleID = rsTeleHead.Fields("ID").Value
		tvTreeView.SelectedNode = tvTreeView.Nodes.Item("T" & CStr(teleID))
		FillTextBox(tvTreeView.SelectedNode)
		
		Exit Sub
TelegramError: 
		TelegramError()
	End Sub
	
	Private Sub ButtonLast()
		' Dim strKey As String    8/2003 efj  removed
		Dim teleID As Short
		
		On Error GoTo TelegramError
		
		If rsTeleHead.BOF And rsTeleHead.EOF Then
			Exit Sub
		End If
		rsTeleHead.MoveLast()
		teleID = rsTeleHead.Fields("ID").Value
		tvTreeView.SelectedNode = tvTreeView.Nodes.Item("T" & CStr(teleID))
		FillTextBox(tvTreeView.SelectedNode)
		
		Exit Sub
TelegramError: 
		TelegramError()
	End Sub
	
	Private Sub ButtonDelete()
		Dim strKey As String
		Dim teleID As Short
		
		On Error GoTo TelegramError
		
		strKey = tvTreeView.SelectedNode.Name
		If VB.Left(strKey, 1) <> "T" Then
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: MSComctlLib.Node property tvTreeView.SelectedItem.FirstSibling was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If tvTreeView.SelectedNode.Name = tvTreeView.SelectedNode.FirstSibling.Name Then
			tvTreeView.SelectedNode = tvTreeView.SelectedNode.NextNode
		Else
			tvTreeView.SelectedNode = tvTreeView.SelectedNode.PrevNode
		End If
		
		teleID = CShort(Mid(strKey, 2))
		'UPGRADE_WARNING: MSComctlLib.Nodes method tvTreeView.Nodes.Remove has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		tvTreeView.Nodes.RemoveAt(strKey)
		
		'Delete the records
		rsTeleHead.Index = "PrimaryKey"
		rsTeleHead.Seek("=", teleID)
		If Not rsTeleHead.NoMatch Then
			rsTeleHead.Delete()
			rsTeleBody.Index = "PrimaryKey"
			rsTeleBody.Seek("=", teleID)
			While Not rsTeleBody.NoMatch
				rsTeleBody.Delete()
				rsTeleBody.Seek("=", teleID)
			End While
		End If
		
		'rsTeleHead.MoveNext
		'Now move to the next node
		'If rsTeleHead.EOF Then
		'    If Not rsTeleHead.BOF Then
		'        rsTeleHead.MovePrevious
		'    Else
		'        Exit Sub
		'    End If
		'End If
		
		'Fill the box in the new mode
		If rsTeleHead.BOF And rsTeleHead.EOF Then
			txtBody.Text = vbNullString
		Else
			'teleID = rsTeleHead.Fields("ID")
			'Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
			FillTextBox(tvTreeView.SelectedNode)
		End If
		
		Exit Sub
		
TelegramError: 
		TelegramError()
	End Sub
	
	Private Sub ButtonNew(ByRef MessageType As eMessageCreate)
		Dim strKey As String
		Dim teleID As Short
		Dim X As Short
		' Dim ReplyNation As String    8/2003 efj  removed
		
		'Fill the combo box
		Nations.FillListBox(lbNation)
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			lbNation.Items.Insert(0, New VB6.ListBoxItem("The Deity", 0))
		End If
		
		'Set to telegram
		Option1.Checked = True
		
		'Set controls
		chkSave.Visible = True
		lbNation.Visible = True
		labNation.Visible = False
		
		'Find the current telegram
		SelectedTelegram = 0
		If MessageType = eMessageCreate.MESS_REPLY Or MessageType = eMessageCreate.MESS_FORWARD Then
			'if no telegram is selected, quit
			On Error GoTo ButtonNew_Exit
			strKey = tvTreeView.SelectedNode.Name
			On Error GoTo 0
			If VB.Left(strKey, 1) = "T" Then
				teleID = CShort(Mid(strKey, 2))
				rsTeleHead.Seek("=", teleID)
				If Not rsTeleHead.NoMatch Then
					For X = 0 To lbNation.Items.Count - 1
						If MessageType = eMessageCreate.MESS_REPLY And VB6.GetItemData(lbNation, X) = CDbl(CStr(rsTeleHead.Fields("From").Value)) Then
							lbNation.SetSelected(X, True)
						End If
					Next X
					chkInclude.Visible = True
					SelectedTelegram = teleID
					chkInclude.CheckState = System.Windows.Forms.CheckState.Checked
				End If
			End If
		End If
		tvTreeView.Visible = False
		
		NewTelegram = True
		txtBody.Text = vbNullString
		txtBody.ReadOnly = False
		
		'Set buttons
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_FIRST).Enabled = False 'All back
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_BACK).Enabled = False 'Back 0ne
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_NEXT).Enabled = False 'forward 0ne
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_LATEST).Enabled = False 'All forward
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_NEW).Enabled = False 'New Tele
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_REPLY).Enabled = False 'Reply
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_FORWARD).Enabled = False 'Forward
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_IMPORT).Enabled = False 'Import Intelligence 103003 rjk: Added
		
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_DELETE).Enabled = True 'Delete
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_SEND).Enabled = True 'Send
		
		'Set titles
		lblTitle(0).Text = "Send Telegram To:"
		
		'Fill in text if its a reply or a forward
		If MessageType = eMessageCreate.MESS_FORWARD Then
			bForwardMessage = True
		Else
			bForwardMessage = False
		End If
		If MessageType = eMessageCreate.MESS_REPLY Or MessageType = eMessageCreate.MESS_FORWARD Then
			chkInclude_CheckStateChanged(chkInclude, New System.EventArgs())
		End If
		
ButtonNew_Exit: 
	End Sub
	
	Private Sub ButtonCancel()
		' Dim strKey As String    8/2003 efj  removed
		' Dim teleID As Integer    8/2003 efj  removed
		
		'Set controls
		lbNation.Visible = False
		labNation.Visible = False
		chkSave.Visible = False
		chkInclude.Visible = False
		tvTreeView.Visible = True
		NewTelegram = False
		txtBody.Text = vbNullString
		txtBody.ReadOnly = True
		FillTextBox(tvTreeView.SelectedNode)
		lblTitle(0).Text = "Received Telegrams"
		'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		Timer1.Interval = 0
	End Sub
	
	Private Sub ButtonSend()
		Dim n As Short
		Dim strTel As String
		Dim hldNatNumber As Short
		Dim hldNation As String
		Dim X As Short
		
		Dim strTemp As String
		Dim strSource As String
		Dim strTarget() As String
		Dim numTelegrams As Short
		
		'Must have a nation !
		If lbNation.SelectedItems.Count = 0 And Option1.Checked Then
			MsgBox("Must enter a Nation to send a telegram.",  , "Error")
			Exit Sub
		End If
		
		If chkSave.CheckState = System.Windows.Forms.CheckState.Checked And Option1.Checked Then
			hldNatNumber = -1
			For X = 0 To lbNation.Items.Count - 1
				If lbNation.GetSelected(X) Then
					hldNatNumber = VB6.GetItemData(lbNation, X)
					hldNation = VB6.GetItemString(lbNation, X)
				End If
			Next X
			
			If hldNatNumber < 0 Then 'still couldn't figure out number, so abort
				chkSave.CheckState = System.Windows.Forms.CheckState.Unchecked
			Else
				IncomingTelegramLine("> Telegram To " & hldNation & " (#" & CStr(hldNatNumber) & ")" & " dated " & VB6.Format(DateAdd(Microsoft.VisualBasic.DateInterval.Minute, rsVersion.Fields("Time Zone Adj").Value, Now), "ddd mmm dd hh:mm:ss yyyy"), True)
			End If
			
		End If
		
		''fixme! drk@unxsoft.com 2/26/03 don't want to generate telegrams with lines over 80 chars
		strSource = txtBody.Text
		'UPGRADE_WARNING: Lower bound of array strTarget was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim strTarget(1)
		numTelegrams = 1
		Do 
			n = InStr(strSource, vbLf)
			If n = 0 Then n = Len(strSource) + 1
			strTemp = VB.Left(strSource, n - 1)
			strSource = Mid(strSource, n + 1)
			
			'remove carrage returns ''why? drk@unxsoft.com 2/26
			n = InStr(strTemp, vbCr)
			While n > 0
				strTemp = VB.Left(strTemp, n - 1) & Mid(strTemp, n + 1)
				n = InStr(strTemp, vbCr)
			End While
			
			If chkSave.CheckState = System.Windows.Forms.CheckState.Checked And Option1.Checked Then
				IncomingTelegramLine(strTemp, False)
			End If
			
			'just in case we get a real long line, split it up
			'110303 rjk: Modified to account for the server not support up 1024 characters,
			'            it only supports 1022 characters in a telegram line.
			While Len(strTemp) > 1021
				n = 1021 - Len(strTarget(numTelegrams))
				strTarget(numTelegrams) = strTarget(numTelegrams) & VB.Left(strTemp, n)
				strTemp = Mid(strTemp, n + 1)
				numTelegrams = numTelegrams + 1 '110303 rjk: Fixed to support lines longer than 2044
				'UPGRADE_WARNING: Lower bound of array strTarget was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve strTarget(numTelegrams)
			End While
			
			If (Len(strTarget(numTelegrams)) + Len(strTemp)) >= 1022 Then
				numTelegrams = numTelegrams + 1
				'UPGRADE_WARNING: Lower bound of array strTarget was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim Preserve strTarget(numTelegrams)
			End If
			
			strTarget(numTelegrams) = strTarget(numTelegrams) & strTemp & vbLf
			
		Loop While Len(strSource) > 0
		
		'run thru multiple telegrams if necessay
		For n = 1 To numTelegrams
			frmEmpCmd.SubmitEmpireCommand("te1 " & strTarget(n), False)
			
			If Option1.Checked Then
				'Build recipient string
				strTel = vbNullString
				For X = 0 To lbNation.Items.Count - 1
					If lbNation.GetSelected(X) Then
						strTel = strTel & CStr(VB6.GetItemData(lbNation, X)) & " "
					End If
				Next X
				frmEmpCmd.SubmitEmpireCommand("telegram " & strTel, False)
			ElseIf Option2.Checked Then 
				frmEmpCmd.SubmitEmpireCommand("announce ", False)
			ElseIf Option3.Checked Then 
				'Build recipient string
				strTel = vbNullString
				For X = 0 To lbNation.Items.Count - 1
					If lbNation.GetSelected(X) Then
						strTel = CStr(VB6.GetItemData(lbNation, X))
					End If
				Next X
				frmEmpCmd.SubmitEmpireCommand("flash " & strTel, False)
			ElseIf optMotd.Checked Then 
				frmEmpCmd.SubmitEmpireCommand("turn mess ", False)
			End If
		Next n
		
		'Set controls
		lbNation.Visible = False
		labNation.Visible = False
		chkSave.Visible = False
		chkInclude.Visible = False
		tvTreeView.Visible = True
		NewTelegram = False
		txtBody.Text = vbNullString
		txtBody.ReadOnly = True
		FillTextBox(tvTreeView.SelectedNode)
		lblTitle(0).Text = "Received Telegrams"
		'UPGRADE_WARNING: Timer property Timer1.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
		Timer1.Interval = 0
		Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
	End Sub
	
	'103003 rjk: Added Import Intelligence
	Private Sub ButtonImport()
		Dim strKey As String
		Dim teleID As Short
		Dim strTelegram As String
		
		'Find the current telegram
		SelectedTelegram = 0
		On Error GoTo ButtonImport_Exit
		strKey = tvTreeView.SelectedNode.Name
		On Error GoTo 0
		
		strTelegram = vbNullString
		
		If VB.Left(strKey, 1) = "T" Then
			teleID = CShort(Mid(strKey, 2))
			rsTeleHead.Seek("=", teleID)
			If Not rsTeleHead.NoMatch Then
				rsTeleBody.Seek("=", teleID, 1)
				If Not rsTeleBody.NoMatch Then
					Do While (Not rsTeleBody.EOF)
						If rsTeleBody.Fields("ID").Value <> teleID Then
							Exit Do
						End If
						If Len(rsTeleBody.Fields("Body").Value) > 0 Then
							strTelegram = strTelegram + rsTeleBody.Fields("Body").Value + vbLf
						End If
						rsTeleBody.MoveNext()
					Loop 
				End If
			End If
		End If
		
		ImportIntelligenceOffset(strTelegram)
ButtonImport_Exit: 
	End Sub
	
	Private Sub LoadTree()
		Dim title As String
		Dim nodX As System.Windows.Forms.TreeNode
		' Dim nodP As Node    8/2003 efj  removed
		Dim X As Short
		Dim c As Short
		Dim ntele As Short
		Dim nteleexp As Short
		
		On Error Resume Next
		
		'Clear out old nodes
		While tvTreeView.Nodes.Count > 0
			'UPGRADE_WARNING: MSComctlLib.Nodes method tvTreeView.Nodes.Remove has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			tvTreeView.Nodes.RemoveAt(1)
		End While
		
		nodX = tvTreeView.Nodes.Add("R", "Telegrams")
		
		'Set the class nodes
		For X = TELEGRAM_CLASS_INBOX + 1 To NUMBER_OF_TELEGRAM_CLASSES
			Select Case X
				Case TELEGRAM_CLASS_INBOX
					title = "In Box"
				Case TELEGRAM_CLASS_COMBAT
					title = "Combat Reports"
				Case TELEGRAM_CLASS_BULLETIN
					title = "Bulletins"
				Case TELEGRAM_CLASS_PLAYER
					title = "Player Telegrams"
				Case TELEGRAM_CLASS_INTELLIGENCE '103003 rjk: Added Import Intelligence
					title = "Intelligence Reports"
				Case TELEGRAM_CLASS_PRODUCTION
					title = "Production Reports"
				Case TELEGRAM_CLASS_REPORTS
					title = "Saved Reports"
				Case TELEGRAM_CLASS_GENERAL
					title = "Misc."
				Case Else
					title = "Other"
			End Select
			'UPGRADE_WARNING: Add method behavior has changed Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DBD08912-7C17-401D-9BE9-BA85E7772B99"'
			nodX = tvTreeView.Nodes.Find("R", True)(0).Nodes.Add("C" & CStr(X), title)
		Next X
		nodX.EnsureVisible()
		
		ntele = 0
		nteleexp = 0
		
		Dim hldIndex As String
		hldIndex = rsTeleHead.Index
		rsTeleHead.Index = "Received"
		If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
			rsTeleHead.MoveLast()
			While Not rsTeleHead.BOF
				ntele = ntele + 1
				c = rsTeleHead.Fields("Class").Value
				If rsTeleHead.Fields("Exported").Value <> "Y" Then
					nteleexp = nteleexp + 1
				End If
				
				'Each player gets a subclass
				If c = TELEGRAM_CLASS_PLAYER Then
					X = rsTeleHead.Fields("From").Value
					AddPlayerNode(X, "P")
					'UPGRADE_WARNING: Add method behavior has changed Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DBD08912-7C17-401D-9BE9-BA85E7772B99"'
					nodX = tvTreeView.Nodes.Find("P" & CStr(X), True)(0).Nodes.Add("T" & CStr(rsTeleHead.Fields("ID").Value), rsTeleHead.Fields("Subject").Value)
				ElseIf c = TELEGRAM_CLASS_INTELLIGENCE Then  '103003 rjk: Added Intelligence Reports
					X = rsTeleHead.Fields("From").Value
					AddPlayerNode(X, "I")
					'UPGRADE_WARNING: Add method behavior has changed Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DBD08912-7C17-401D-9BE9-BA85E7772B99"'
					nodX = tvTreeView.Nodes.Find("I" & CStr(X), True)(0).Nodes.Add("T" & CStr(rsTeleHead.Fields("ID").Value), rsTeleHead.Fields("Subject").Value)
				Else
					'UPGRADE_WARNING: Add method behavior has changed Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DBD08912-7C17-401D-9BE9-BA85E7772B99"'
					nodX = tvTreeView.Nodes.Find("C" & CStr(c), True)(0).Nodes.Add("T" & CStr(rsTeleHead.Fields("ID").Value), rsTeleHead.Fields("Subject").Value)
				End If
				rsTeleHead.MovePrevious()
			End While
			rsTeleHead.MoveLast()
		End If
		
		rsTeleHead.Index = hldIndex
		
		'UPGRADE_ISSUE: Constant tvwTreelinesText was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: MSComctlLib.TreeView property tvTreeView.Style was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		tvTreeView.Style = tvwTreelinesText ' Style 4.
		tvTreeView.BorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		'UPGRADE_WARNING: Lower bound of collection sbStatusBar.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sbStatusBar.Items.Item(1).Text = CStr(ntele) & " Telegrams, "
		'UPGRADE_WARNING: Lower bound of collection sbStatusBar.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		sbStatusBar.Items.Item(2).Text = CStr(nteleexp) & " not exported"
		TreeDirty = False
		lblTitle(1).Text = vbNullString
	End Sub
	
	'103103 rjk: Added strClass to separate the Intelligence Class from Player Class
	Private Sub AddPlayerNode(ByRef X As Short, ByRef strClass As String)
		Dim Node As System.Windows.Forms.TreeNode
		Dim PlayerName As String
		
		On Error GoTo PlayerNode_New_Node
		Node = tvTreeView.Nodes.Item(strClass & CStr(X))
		Exit Sub
		
PlayerNode_New_Node: 
		On Error Resume Next
		'Create a new node
		PlayerName = Nations.NationName(X)
		If Len(PlayerName) = 0 Then
			PlayerName = "Nation #" & CStr(X)
		End If
		
		'UPGRADE_WARNING: Add method behavior has changed Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DBD08912-7C17-401D-9BE9-BA85E7772B99"'
		Node = tvTreeView.Nodes.Find("C" & CStr(rsTeleHead.Fields("Class").Value), True)(0).Nodes.Add(strClass & CStr(X), PlayerName)
	End Sub
	
	Private Sub tvTreeView_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvTreeView.NodeMouseClick
		Dim Node As System.Windows.Forms.TreeNode = eventArgs.Node
		'Put the selected telegram in the text box
		FillTextBox(Node)
	End Sub
	
	Private Sub FillTextBox(ByVal Node As System.Windows.Forms.TreeNode)
		Dim teleID As Short
		
		On Error Resume Next
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_FIRST).Enabled = True 'All back
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_BACK).Enabled = False 'Back 0ne
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_NEXT).Enabled = False 'forward 0ne
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_LATEST).Enabled = True 'All forward
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_NEW).Enabled = True 'New Tele
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_REPLY).Enabled = False 'Reply
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_FORWARD).Enabled = False 'Forward
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_IMPORT).Enabled = False 'Import Intelligence 103003 rjk: Added
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_DELETE).Enabled = False 'Delete
		'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		tbToolBar.Items.Item(BUTTON_SEND).Enabled = False 'Send
		
		
		' Dim hldFont As String    8/2003 efj  removed
		Dim hldText As String
		
		txtBody.Text = vbNullString
		hldText = vbNullString
		If VB.Left(Node.Name, 1) = "T" Then
			teleID = CShort(Mid(Node.Name, 2))
			rsTeleBody.Index = "PrimaryKey"
			rsTeleBody.Seek("=", teleID, 1)
			
			'Load telegram body into text box
			If Not rsTeleBody.NoMatch Then
				CurTele = teleID
				Do While (Not rsTeleBody.EOF)
					If rsTeleBody.Fields("ID").Value <> teleID Then
						Exit Do
					End If
					'0901903 rjk: removed linefeed for the reports, as data is not save on per line
					' bases anymore but in 255 byte chucks.  However, bulletins are still saved one line at
					' time with no linefeeds.
					'101503 rjk: If the record is full then no new line is required.
					'            Blank added at insertion time.
					If Len(rsTeleBody.Fields("Body").Value) = 255 Then
						'If InStr(rsTeleBody.Fields("Body"), vbLf) > 0 Then 'Reports
						hldText = hldText + rsTeleBody.Fields("Body").Value
					ElseIf VB.Right(rsTeleBody.Fields("Body").Value, 1) <> vbLf Then 
						hldText = hldText + rsTeleBody.Fields("Body").Value + vbLf
					Else
						hldText = hldText + rsTeleBody.Fields("Body").Value
					End If
					'txtBody.Text = txtBody.Text + rsTeleBody.Fields("Body") + vbLf
					rsTeleBody.MoveNext()
				Loop 
			End If
			txtBody.Text = hldText
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(BUTTON_BACK).Enabled = True 'Back 0ne
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(BUTTON_NEXT).Enabled = True 'forward 0ne
			'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			tbToolBar.Items.Item(BUTTON_DELETE).Enabled = True 'Delete
			
			'Set title
			rsTeleHead.Index = "PrimaryKey"
			rsTeleHead.Seek("=", CShort(Mid(Node.Name, 2)))
			If Not rsTeleHead.NoMatch Then
				lblTitle(1).Text = rsTeleHead.Fields("Title").Value
				If rsTeleHead.Fields("Read").Value <> "Y" Then
					rsTeleHead.Edit()
					rsTeleHead.Fields("Read").Value = "Y"
					rsTeleHead.Update()
				End If
				If rsTeleHead.Fields("Class").Value = TELEGRAM_CLASS_PLAYER Then
					'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					tbToolBar.Items.Item(BUTTON_REPLY).Enabled = True 'Reply
				End If
				If rsTeleHead.Fields("Class").Value = TELEGRAM_CLASS_INTELLIGENCE Then
					'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					tbToolBar.Items.Item(BUTTON_IMPORT).Enabled = True 'Import Intelligence 103003 rjk: Added
				End If
				'UPGRADE_WARNING: Lower bound of collection tbToolBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				tbToolBar.Items.Item(BUTTON_FORWARD).Enabled = True 'Forward
			End If
		Else
			lblTitle(1).Text = vbNullString
		End If
	End Sub
	
	'112903 rjk: Remove the question on clearing the telegrams as there is a separate form now.
	Public Sub ClearTelegrams()
		'if no telegrams, then quit
		If rsTeleHead.BOF And rsTeleHead.EOF Then
			Exit Sub
		End If
		
		DeleteAllRecords(rsTeleHead)
		DeleteAllRecords(rsTeleBody)
		
		LoadTree() '111403 rjk: Clear Telegram window after removing all the telegrams.
	End Sub
	
	Private Sub txtBody_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBody.TextChanged
		If Not txtBody.ReadOnly Then
			'UPGRADE_WARNING: Lower bound of collection sbStatusBar.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			sbStatusBar.Items.Item(2).Text = CStr(Len(txtBody.Text)) & " characters"
			ColorLength()
		End If
	End Sub
	
	Private Sub ColorLength()
		Dim lIndex As Integer
		Dim lCharCount As Integer
		Dim lStartPos As Integer
		Dim lCharPos As Integer
		Dim lLength As Integer
		Dim iTelegramTypeIndex As Short
		Dim iSoundIndex As Short
		
		If Option1.Checked = True Then
			iTelegramTypeIndex = WinAceRoutines.enumTelegramOptionType.totTelegram
		ElseIf Option2.Checked = True Then 
			iTelegramTypeIndex = WinAceRoutines.enumTelegramOptionType.totAnnouncement
		ElseIf Option3.Checked = True Then 
			iTelegramTypeIndex = WinAceRoutines.enumTelegramOptionType.totFlash
		ElseIf optMotd.Checked = True Then 
			iTelegramTypeIndex = WinAceRoutines.enumTelegramOptionType.totMOTD
		Else
			Exit Sub
		End If
		
		lCharCount = 0
		lStartPos = 0
		lLength = txtBody.SelectionLength
		lCharPos = txtBody.SelectionStart
		If Not tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).bEnabled And Not tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).bEnabled Then
			txtBody.SelectionStart = 0
			txtBody.SelectionLength = Len(txtBody.Text)
			txtBody.SelectionColor = System.Drawing.Color.Black
			txtBody.SelectionStart = lCharPos
			txtBody.SelectionLength = lLength
			If Not tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlHardLimit) Then
				Exit Sub
			End If
		End If
		For lIndex = 0 To Len(txtBody.Text) - 1
			lCharCount = lCharCount + 1
			If tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlHardLimit).bEnabled And lCharCount = tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlHardLimit).iColumn + 1 And Mid(txtBody.Text, lIndex + 1, 1) <> vbCr Then
				txtBody.Text = VB.Left(txtBody.Text, lIndex) & vbCrLf & VB.Right(txtBody.Text, Len(txtBody.Text) - lIndex)
				If lCharPos > lIndex Then
					lCharPos = lCharPos + 2 'Advance the current point past the vbCrLf
				End If
			End If
			If Mid(txtBody.Text, lIndex + 1, 1) = vbCr Then
				txtBody.SelectionStart = lStartPos
				txtBody.SelectionLength = lIndex - lStartPos + 1
				If tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).bEnabled And lCharCount > tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).iColumn Then
					txtBody.SelectionColor = System.Drawing.Color.Red
				ElseIf tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).bEnabled And lCharCount > tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).iColumn Then 
					txtBody.SelectionColor = System.Drawing.Color.Blue
				ElseIf lIndex = Len(txtBody.Text) - 1 Then 
					txtBody.SelectionColor = System.Drawing.Color.Black
				End If
				lCharCount = 0
				lStartPos = lIndex + 1
			ElseIf Mid(txtBody.Text, lIndex + 1, 1) = vbLf Then 
				lCharCount = 0
				lStartPos = lIndex + 1
			ElseIf lIndex = Len(txtBody.Text) - 1 Then 
				txtBody.SelectionStart = lStartPos
				txtBody.SelectionLength = lIndex - lStartPos + 1
				If tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).bEnabled And lCharCount > tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).iColumn Then
					txtBody.SelectionColor = System.Drawing.Color.Red
					If lCharCount = tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).iColumn + 1 Then
						For iSoundIndex = 1 To tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).eTelegramSound
							Beep()
						Next iSoundIndex
					End If
				ElseIf tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).bEnabled And lCharCount > tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).iColumn Then 
					txtBody.SelectionColor = System.Drawing.Color.Blue
					If lCharCount = tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).iColumn + 1 Then
						For iSoundIndex = 1 To tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).eTelegramSound
							Beep()
						Next iSoundIndex
					End If
				ElseIf lIndex = Len(txtBody.Text) - 1 Then 
					txtBody.SelectionColor = System.Drawing.Color.Black
				End If
			ElseIf tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).bEnabled And lCharCount = tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).iColumn Then 
				txtBody.SelectionStart = lStartPos
				txtBody.SelectionLength = lIndex - lStartPos + 1
				txtBody.SelectionColor = System.Drawing.Color.Black
				lStartPos = lIndex + 1
			ElseIf tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).bEnabled And lCharCount = tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlSoftLimit).iColumn Then 
				txtBody.SelectionStart = lStartPos
				txtBody.SelectionLength = lIndex - lStartPos + 1
				If Not tTelegramOptions(iTelegramTypeIndex, WinAceRoutines.enumTelegramLevel.tlWarning).bEnabled Then
					txtBody.SelectionColor = System.Drawing.Color.Black
				Else
					txtBody.SelectionColor = System.Drawing.Color.Blue
				End If
				lStartPos = lIndex + 1
			End If
		Next lIndex
		txtBody.SelectionStart = lCharPos
		txtBody.SelectionLength = lLength
	End Sub
	
	Private Sub SetRecipientLabel()
		Dim X As Short
		Dim strNationList As String
		Dim strTemp As String
		
		'Look for what is selected
		For X = 0 To lbNation.Items.Count - 1
			If lbNation.GetSelected(X) Then
				strNationList = strNationList & VB6.GetItemString(lbNation, X) & ", "
			End If
		Next X
		
		'Now set label
		If Option1.Checked Then
			strTemp = "Telegram"
		Else
			strTemp = "Flash"
		End If
		
		If Len(strNationList) = 0 Then
			lblTitle(1).Text = "Enter New " & strTemp & ", Select Recipient"
		Else
			lblTitle(1).Text = strTemp & " For: " & VB.Left(strNationList, Len(strNationList) - 2)
		End If
		
	End Sub
	
	Private Sub TelegramError()
		MsgBox("ERROR - Your telegram database is corrupted.  WinACE was not able" & " to move to the requested Telegram.  You should take Export, then Clear, then Import off" & " the File Menu to repair the database", MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Database Error")
	End Sub
End Class