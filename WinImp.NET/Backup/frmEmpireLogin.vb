Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmEmpireLogin
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: Removed dead variables
	'101603 rjk: Added offline mode during startup
	'            Storing and Retrieving of the AutoUpdate and AutoRead for each game
	'240104 rjk: Fixed retry.  Also corrected the spelling of successfullconnnection.
	'200304 rjk: Detect running out of BTU's and logout and back in.
	'250405 rjk: Close the socket when the remote end is closing the socket in
	'            CheckLostConnection.
	'260505 rjk: Switched UTF8 from an registry option to login option.
	'160705 rjk: Added http (XmlRpc) proxy for access from behind a firewall.
	'200506 rjk: Switch XMLRPC to asychronous communications.
	'            Fixed missing sets and xmlrpcServer object checks.
	'190606 rjk: Added Theme Game control, duplicate of the option Theme Game tab.
	'090706 rjk: Added HeavyMetal Theme Game.
	
	Public ConnectedtoHost As Boolean
	Public bOffline As Boolean '092603 rjk: Added to work offline
	
	'UPGRADE_WARNING: Event chkHTTP.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkHTTP_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkHTTP.CheckStateChanged
		If chkHTTP.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			cmdHTTP.Enabled = True
		Else
			cmdHTTP.Enabled = False
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error Resume Next
		
		'set the global var to false
		'to denote a failed login
		If bXMLRPC Then
			xmlrpcServer.CloseConnection()
			'UPGRADE_NOTE: Object xmlrpcServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			xmlrpcServer = Nothing
		Else
			frmEmpCmd.Winsock.Close()
		End If
		ConnectedtoHost = False
		Me.Hide()
	End Sub
	
	Private Sub cmdConnect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConnect.Click
		Dim filenumber As Short
		'Dim strDeity As Integer 101603 Switched to string based on usage
		Dim strDeity As String
		'101603 rjk: Added to store the current setting of AutoUpdate and AutoRead
		Dim strAutoUpdate As String
		Dim strAutoRead As String
		Dim strUTF8 As String
		
		'Check that all data is entered
		If Len(txtHostName.Text) = 0 Or Len(txtPort.Text) = 0 Or (Len(txtUserName.Text) = 0 And chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked) Or (Len(txtUser.Text) = 0 And chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked) Or (Len(txtPassword.Text) = 0 And chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked) Then
			Label1.Text = "Error: Must fill each field"
			Exit Sub
		End If
		
		'Set Deity Flag
		If chkDeity.CheckState = System.Windows.Forms.CheckState.Checked Then
			strDeity = "1"
		Else
			strDeity = "0"
		End If
		
		'101603 rjk: Creating string to store current setting of AutoUpdate and AutoRead
		If chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked Then
			strAutoUpdate = "1"
		Else
			strAutoUpdate = "0"
		End If
		If chkAutoRead.CheckState = System.Windows.Forms.CheckState.Checked Then
			strAutoRead = "1"
		Else
			strAutoRead = "0"
		End If
		If chkUTF8.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			strUTF8 = "1"
		Else
			strUTF8 = "0"
		End If
		
		If Me.chkHTTP.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			bXMLRPC = True
		Else
			bXMLRPC = False
		End If
		frmHTTPSettings.SaveXMLRPCParameters()
		
		SaveSetting(APPNAME, "Options", "Star Wars", CStr(False))
		SaveSetting(APPNAME, "Options", "Retro", CStr(False))
		SaveSetting(APPNAME, "Options", "2K4", CStr(False))
		SaveSetting(APPNAME, "Options", "EE9", CStr(False))
		SaveSetting(APPNAME, "Options", "SP2005", CStr(False))
		SaveSetting(APPNAME, "Options", "SPAtlantis", CStr(False))
		SaveSetting(APPNAME, "Options", "HeavyMetal", CStr(False))
		Select Case cmbThemeGame.SelectedIndex
			Case 1
				SaveSetting(APPNAME, "Options", "Star Wars", CStr(True))
			Case 2
				SaveSetting(APPNAME, "Options", "Retro", CStr(True))
			Case 3
				SaveSetting(APPNAME, "Options", "2K4", CStr(True))
			Case 4
				SaveSetting(APPNAME, "Options", "EE9", CStr(True))
			Case 5
				SaveSetting(APPNAME, "Options", "SP2005", CStr(True))
			Case 6
				SaveSetting(APPNAME, "Options", "SPAtlantis", CStr(True))
			Case 7
				SaveSetting(APPNAME, "Options", "HeavyMetal", CStr(True))
		End Select
		
		'Save connection data
		filenumber = FreeFile ' Get unused file number.
		If VerifySubDirectory("Games", True) Then
			FileOpen(filenumber, My.Application.Info.DirectoryPath & "\Games\" & GameName & ".egp", OpenMode.Output)
			'101603 rjk: Added storing current setting of AutoUpdate and AutoRead
			WriteLine(filenumber, Trim(txtHostName.Text), Trim(txtPort.Text), txtUserName.Text, txtPassword.Text, txtUser.Text, strDeity, strAutoUpdate, strAutoRead, strUTF8)
			FileClose(filenumber)
		End If
		
		On Error GoTo Connect_Error
		
		'Make the connection to the Empire Server
		If Not ConnectedtoHost Then
			If bXMLRPC Then
				On Error GoTo XMLRPC_Error
				xmlrpcServer = New EmpXMLRPC
				xmlrpcServer.OpenConnection(Trim(txtHostName.Text), Val(txtPort.Text))
			Else
				frmEmpCmd.Winsock.RemoteHost = Trim(txtHostName.Text)
				frmEmpCmd.Winsock.RemotePort = Val(txtPort.Text)
				frmEmpCmd.Winsock.Connect()
			End If
			
			'Set the timer to check for the Winsock State
			Timer1.Interval = 350
			Timer1.Enabled = True
			Timer1_Tick(Timer1, New System.EventArgs())
		Else
			Me.Hide() 'Your done - must be reconnect.
		End If
		Exit Sub
		
Connect_Error: 
		Label1.Text = "Error: " & Err.Description
		Exit Sub
		
XMLRPC_Error: 
		Label1.Text = "Error: " & Err.Description
		MsgBox("You need to load Microsoft XML Parser 4.0 before you can HTTP - XML RPC capability", MsgBoxStyle.OKOnly)
	End Sub
	
	Public Sub go()
		Dim filenumber As Integer
		Dim s1 As String
		Dim s2 As String
		Dim s3 As String
		Dim s4 As String
		
		frmHTTPSettings.LoadXMLRPCParameters()
		If bXMLRPC Then
			chkHTTP.CheckState = System.Windows.Forms.CheckState.Checked
			cmdHTTP.Enabled = True
		Else
			chkHTTP.CheckState = System.Windows.Forms.CheckState.Unchecked
			cmdHTTP.Enabled = False
		End If
		'Check for game name file - if not found, then must be a new entry
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(My.Application.Info.DirectoryPath & "\Games\" & GameName & ".egp")) > 0 Then
			'Load the parameters from the current file
			filenumber = FreeFile ' Get unused file number.
			FileOpen(filenumber, My.Application.Info.DirectoryPath & "\Games\" & GameName & ".egp", OpenMode.Input)
			Input(filenumber, s1)
			Input(filenumber, s2)
			Input(filenumber, s3)
			Input(filenumber, s4)
			txtHostName.Text = s1
			txtPort.Text = s2
			txtUserName.Text = s3
			txtPassword.Text = s4
			On Error Resume Next
			Input(filenumber, s1)
			txtUser.Text = s1
			Input(filenumber, s1)
			If s1 = "1" Then
				chkDeity.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkDeity.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
			'101603 rjk: Retrieving of the AutoUpdate and AutoRead for each game
			If Not EOF(filenumber) Then
				Input(filenumber, s1)
				Input(filenumber, s2)
				If s1 = "1" Then
					chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					chkUpdate.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				If s2 = "1" Then
					chkAutoRead.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					chkAutoRead.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
			End If
			If Not EOF(filenumber) Then
				Input(filenumber, s1)
				If s1 = "1" Then
					chkUTF8.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					chkUTF8.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
			End If
			FileClose(filenumber)
		End If
		
		cmbThemeGame.Items.Add("None")
		cmbThemeGame.Items.Add("Star Wars")
		cmbThemeGame.Items.Add("Retro")
		cmbThemeGame.Items.Add("2K4")
		cmbThemeGame.Items.Add("EE9")
		cmbThemeGame.Items.Add("SP2005")
		cmbThemeGame.Items.Add("SPAtlantis")
		cmbThemeGame.Items.Add("HeavyMetal")
		
		If CBool(GetSetting(APPNAME, "Options", "Star Wars", CStr(False))) Then
			cmbThemeGame.SelectedIndex = 1
		ElseIf CBool(GetSetting(APPNAME, "Options", "Retro", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 2
		ElseIf CBool(GetSetting(APPNAME, "Options", "2K4", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 3
		ElseIf CBool(GetSetting(APPNAME, "Options", "EE9", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 4
		ElseIf CBool(GetSetting(APPNAME, "Options", "SP2005", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 5
		ElseIf CBool(GetSetting(APPNAME, "Options", "SPAtlantis", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 6
		ElseIf CBool(GetSetting(APPNAME, "Options", "HeavyMetal", CStr(False))) Then 
			cmbThemeGame.SelectedIndex = 7
		Else
			cmbThemeGame.SelectedIndex = 0
		End If
		
		lblGame.Text = "Game: " & GameName
	End Sub
	
	Private Sub cmdHTTP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHTTP.Click
		frmHTTPSettings.ShowDialog()
	End Sub
	
	Private Sub cmdOffline_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOffline.Click
		bOffline = True
		cmdConnect_Click(cmdConnect, New System.EventArgs())
		Me.Hide()
	End Sub
	
	Private Sub frmEmpireLogin_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		Dim strState As String
		
		If bXMLRPC Then
			If xmlrpcServer Is Nothing Then
				Label1.Text = "Not Connected"
			ElseIf xmlrpcServer.Status Then 
				Label1.Text = "Connected"
			Else
				Label1.Text = "Not Connected"
			End If
			
			'Check that existing connection is OK
			If ConnectedtoHost Then
				If xmlrpcServer.Status Then
					Exit Sub
				Else
					ConnectedtoHost = False
					frmEmpCmd.ConnectedtoHost = False
					frmEmpCmd.CheckLostConnection()
				End If
			End If
			
			If xmlrpcServer Is Nothing Then
				If frmEmpCmd.bRelogin Then
					cmdConnect_Click(cmdConnect, New System.EventArgs())
				End If
			ElseIf xmlrpcServer.Status Then 
				If Not ConnectedtoHost Then
					SuccessfullConnection()
				End If
			Else
				If frmEmpCmd.bRelogin Then
					cmdConnect_Click(cmdConnect, New System.EventArgs())
				End If
			End If
			If xmlrpcServer Is Nothing Then
				Label1.Text = "Not Connected"
			ElseIf xmlrpcServer.Status Then 
				Label1.Text = "Connected"
			Else
				Label1.Text = "Not Connected"
			End If
		Else
			'Check that existing connection is OK
			If ConnectedtoHost Then
				'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
				If frmEmpCmd.Winsock.CtlState = MSWinsockLib.StateConstants.sckConnected Then
					Exit Sub
				Else
					'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
					If frmEmpCmd.Winsock.CtlState = MSWinsockLib.StateConstants.sckClosing Then
						frmEmpCmd.Winsock.Close()
					End If
					ConnectedtoHost = False
					frmEmpCmd.ConnectedtoHost = False
					frmEmpCmd.CheckLostConnection()
				End If
			End If
			
			'Check the state of the Winsock Connection
			'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			Select Case frmEmpCmd.Winsock.CtlState
				Case MSWinsockLib.StateConstants.sckClosed
					strState = "Closed"
					If frmEmpCmd.bRelogin Then
						cmdConnect_Click(cmdConnect, New System.EventArgs())
					End If
				Case MSWinsockLib.StateConstants.sckOpen
					strState = "Open"
				Case MSWinsockLib.StateConstants.sckListening
					strState = "Listening"
				Case MSWinsockLib.StateConstants.sckConnectionPending
					strState = "Connection pending"
				Case MSWinsockLib.StateConstants.sckResolvingHost
					strState = "Resolving Host"
				Case MSWinsockLib.StateConstants.sckHostResolved
					strState = "Host resolved"
				Case MSWinsockLib.StateConstants.sckConnecting
					strState = "Connecting"
				Case MSWinsockLib.StateConstants.sckConnected
					strState = "Connected"
					If Not ConnectedtoHost Then
						SuccessfullConnection()
					End If
				Case MSWinsockLib.StateConstants.sckClosing
					strState = "Peer is closing the connection"
					frmEmpCmd.Winsock.Close()
				Case MSWinsockLib.StateConstants.sckError
					strState = "Error"
					BadConnection()
				Case Else
					'UPGRADE_NOTE: State was upgraded to CtlState. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
					strState = "Unknown - " & CStr(frmEmpCmd.Winsock.CtlState)
			End Select
			
			Label1.Text = strState
		End If
	End Sub
	
	Private Sub SuccessfullConnection()
		ConnectedtoHost = True
		Timer1.Interval = 5000
		Me.Hide()
		If frmEmpCmd.bRetry Or frmEmpCmd.bRelogin Then
			frmEmpCmd.bRelogin = False
			frmEmpCmd.LoginIntoServer()
		End If
	End Sub
	
	Private Sub BadConnection()
		'092603 rjk: Added Cancel to work offline
		'MsgBox "Error - Connection not Established", vbOKOnly, "Connection Error"
		'Timer1.Enabled = False
		'frmEmpCmd.Winsock.Close
		Dim Reply As Short
		
		If bOffline Then
			Exit Sub
		End If
		
		Reply = MsgBox("Error - Connection not Established, use Cancel to work offline", MsgBoxStyle.OKCancel, "Connection Error")
		If Reply <> MsgBoxResult.Cancel Then
			Timer1.Enabled = False
			If bXMLRPC Then
				xmlrpcServer.CloseConnection()
				'UPGRADE_NOTE: Object xmlrpcServer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				xmlrpcServer = Nothing
			Else
				frmEmpCmd.Winsock.Close()
			End If
			bOffline = False
		Else
			bOffline = True
			ConnectedtoHost = False
			frmEmpCmd.ConnectedtoHost = False
			Timer1.Enabled = False
			Me.Hide()
		End If
	End Sub
End Class