Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmEmpCmd
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: Removed dead variables and procedures
	'092603 rjk: Added offline mode
	'092703 rjk: Added changes for offline mode.
	'092903 rjk: Added timestamp ability to ndump.  Added cutoff and level reports.
	'            Fixed ship cargo - wrong command. Added tra(nsport) to nuke command list.
	'            Fixed the maps so they would generate reports.
	'100103 rjk: Added DisplayCmd=False to nation so when enlist or demobilize commands are not bumped by nation update (mil. reserves)
	'100303 rjk: Added treaty report
	'100503 rjk: Removed timestamp ndump
	'100603 rjk: Removed duplicates from nSector list.
	'100703 rjk: Moved Nation list logic into relation report, removes the need for READ_STATE_REQ_REPORT
	'100703 rjk: Added nCountry to detect the change command and submit a nation command to update country name
	'100903 rjk: Added NextCommand to "update aborting command" logic, it prevents a lockup if in a offer treaty command
	'            and an update occurs, but not handle the read telegram correctly.
	'       rjk: Rename Accept to AcceptanceReport
	'101403 rjk: Added DatabaseInProcess to suppress "show sect s" during startup
	'            If the process dump gettings out of step and it will cause run-time errors processing PLANE dump when expecting LAND dump
	'            Changed code to reduce the occurance of run-time errors
	'101503 rjk: Changed processing of lost dumps to prevent run-time errors if the dumps get out of step
	'101503 rjk: Added processing for land sectors in Anti-Sub patrols during recon's.
	'101703 rjk: Removed trade report, as trade is interactive and does not work as a report
	'101703 rjk: Added Strength fields to Sector database
	'101803 rjk: Added Route report, combined duplicate Spy Reports, Added neweff Report
	'102103 rjk: Fixed processing for strength to deal with an empty report.
	'            Updated the processing of update command if it is missing the update time.
	'102503 rjk: Added Shift F9 to forward in the command list
	'102703 rjk: Added detection in ProcessDump for when grid update is required and only update when required.
	'110303 rjk: Change processing of relations to always fill Nations object information not just from database update
	'111603 rjk: Added a check not a legal command in the strength report for sanutary startup.
	'111803 rjk: Change to relations to output as well if the command was submitted from the command line.
	'112003 rjk: Added option to control strength updates
	'112903 rjk: Added additional parameters related to changes for ClearEnemyInfo
	'            Removed MsgBox "WinACE bug!, unsure what Daniel was working on.
	'            Added the ability to generating Hourly Activity Report from the Newspaper
	'120103 rjk: Added the ability to display current players
	'120203 rjk: Switched to frmOptions and boolean options
	'120203 rjk: Fixed a crash when the option clear command box option was toggled and update comes before
	'            any other command is entered.
	'120503 rjk: Skip an invalid llook or look command (Usage)
	'121203 rjk: Added show item.
	'121303 rjk: Added the ability to display headers for the dump commands.
	'121503 rjk: Fixed a problem with relations if the output is not all in one packet (Empire Hub)
	'121703 rjk: Added a check for Diplomatic Relations title to prevent crash with invalid syntax in relations command
	'123103 rjk: Parse an relations change during a np recon
	'240104 rjk: Cleanup retry for dropping connections, automated the kill process. Cleaned up the command matching.
	'012304 rjk: Fixed a you have mail problem.
	'012704 rjk: Fixed shows sector count for strength report, user requested.
	'020504 rjk: Switched (lps)dump to use the headers to process.
	'022004 rjk: Removed extra space from play and kill commands, test server does not work them.
	'060304 rjk: Added Sonar parsing to Navigation form, 080604.
	'080304 rjk: Added Check for St@r W@rs space travel (flying through space)
	'200304 rjk: Detect running out of BTU's and logout and back in.
	'250304 rjk: Fixed ListSanctuary if it is not a blitz.
	'250304 rjk: Trap illegal commands, print error message and continue
	'270304 rjk: Switched over to DeleteAllRecords for clearing a table
	'080404 rjk: Reorganized the automatic response for the attack prompts.
	'120404 rjk: Added check for percent for processing nuke output for sector damage
	'240404 rjk: Fixed so the CountryName and CountryNumber is filled in Offline mode.
	'150504 rjk: Added detection for WinEmpireHub for chat window.
	'150504 rjk: Added processing DUMPS when in WinEmpireHub mode.
	'060604 rjk: Fixed the Flash window to accomodate flash message being split by the server.
	'110604 rjk: Modified LoginUser and NextCommand to be Public.
	'110604 rjk: Added calls to fillin Flash Window.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210704 rjk: Removed Check for St@r W@rs space travel (flying through space)
	'110804 rjk: Added the motd to telegram window for deity mode.
	'150804 rjk: Added support EmpireHub and converted WinEmpireHub.
	'240804 rjk: Fixed 'break' to allow prompt input (changed for 2K4).
	'290804 rjk: Always save telegrams, unless they are a duplicated.
	'151104 rjk: Added the ability to start with your Capital instead of 0,0
	'271104 rjk: Fixed scuttle to show the tradeship results.
	'011204 rjk: Fixed the overflowing of the command buffer (CopyGameDB).
	'271204 rjk: Fixed the output redirection for xdump command.
	'220305 rjk: Added SubmitTelegramRead to single function for requesting telegrams.
	'            Added version check to deal syntax change in 4.2.21.
	'060405 rjk: Enhanced FLASH/BROADCAST to deal with parse errors.
	'260505 rjk: Switched UTF8 from an registry option to login option.
	'030605 rjk: Fixed List Sanctuaries.
	'120705 rjk: Added xdump to commands to be skipped when recording a script file
	'            In SubmitFromCommandLine, only check for updates if the command
	'            is at least 3 characters and bAutoUpdate is enabled.
	'160705 rjk: Added http (XmlRpc) proxy for access from behind a firewall.
	'            Removed blank lines from command line for internal commands.
	'            Suppress the C_INFORM messages.  Suppress telegram notification
	'            when the option is on.  Suppress the update command display when
	'            the option is on.  Added GetOilContent.
	'230705 rjk: Added an option to suppress unit damage reports from shelling.
	'190805 rjk: Clear the Sea table when breaking.
	'200805 rjk: Fixed time parsing of current time. Server 4.2.22 has less spaces.
	'171205 rjk: Fixed Recording Scripts for Empire Hub.
	'281205 rjk: Remove non-functional infrastructure search pattern.
	'010106 rjk: Change Hidden command to Peek to reflect the server change in 4.2.22.
	'180206 rjk: Replace nation with GetNation.  Replace sdump with GetShips.
	'            Replace ldump with GetLandUnits.  Replace ndump with GetNukes.
	'            Replace pdump with GetPlanes.  Replace lost with GetLost.
	'            Replace relation with GetCountryList.
	'            Use xdump to replace mission for 4.3.0.  Use xdump to replace orders for 4.3.0.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'230406 rjk: Added Flash Logging.
	'170506 rjk: Fixed bomb prompt response to deal with SP: Atlantis changes.
	'200506 rjk: Switch XMLRPC to asychronous communications.
	'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
	'            Also SP: Atlantis will have some the xdump nat and coun changes.
	'100606 rjk: Fixed a bug with sailing order results in production report.
	'            The second part of the production report was being lost.
	'            Fixed the break command to request the meta data before requesting
	'            any tables.
	'            Fixed bomb/fly/recon to have an title on the report.  Cleaned up
	'            strReply[12].  Clear the report on startup.
	'160706 rjk: Added AutoView for navigate and march.
	'230706 rjk: Added xdump support for WinEmpireHub.
	'041006 rjk: Fixed automatic viewing for navigating, prevent double 'h' and
	'            allow ship update to work.
	'181006 rjk: Added TTS for flashes.
	'161206 rjk: Automatically detect the change to XDump capable server and
	'            request the meta tables.
	'311206 rjk: Fixed command line torpedo to request a ship update instead of
	'            land update.
	'090607 rjk: Added nav to Land and Plane update lists.
	'180707 rjk: Skip show updates for server 4.3.10
	'090907 rjk: Added code for using xdump game and xdump update to get
	'            next update for server 4.3.10 and newer.
	'120907 rjk: Fixed "time zone adj" for new game starts.
	'291007 rjk: Allow spaces in the user name.
	'220508 rjk: Delete enemy unit when it is blown up.
	'140708 rjk: Added march, attack, assault, navigate to lost list.
	'            Look for items lost from interdiction.
	'160808 rjk: Ignore the 4.3.16 version of show sect c.
	'250309 rjk: Fixed the parsing for the show sect s.
	'151009 rjk: Nav errors to only produce the pop up when the command
	'            is initiatated from the navigate form.  Remove the pop up
	'            command line nav commands.
	'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.
	
	Const EMP_CLIENT_PROTO As Short = 2
	
	Const EMP_C_CMDOK As Short = 0
	Const EMP_C_DATA As Short = 1
	Const EMP_C_INIT As Short = 2
	Const EMP_C_EXIT As Short = 3
	Const EMP_C_FLUSH As Short = 4
	Const EMP_C_NOECHO As Short = 5
	Const EMP_C_PROMPT As Short = 6
	Const EMP_C_REDIR As Short = 8
	Const EMP_C_PIPE As Short = 9
	Const EMP_C_CMDERR As Integer = &HA
	Const EMP_C_BADCMD As Integer = &HB
	Const EMP_C_EXECUTE As Integer = &HC
	Const EMP_C_FLASH As Short = 13
	Const EMP_C_INFORM As Short = 14
	Const EMP_C_HUB_COMMAND As Short = 16 'g'
	Const EMP_C_HUB_PROMPT As Short = 17 'h'
	Const EMP_C_HUB_DATA As Short = 18 'i'
	Const EMP_C_HUB_TOKEN As Short = 21 'l'
	
	Const DUMP_DELAY As Short = 0
	
	Public ConnectedtoHost As Boolean
	Public bRelogin As Boolean
	Public ForceUpdates As Boolean
	Public CmdQ As EmpCmdQueue
	Public CmdinProgress As Boolean
	Public CountryName As String
	Public SubmittedFromCommandLine As Boolean
	Public bAutoUpdate As Boolean
	Public bAutoRead As Boolean
	Public bRetry As Boolean
	Public ExpectSpecificResponse As Short
	Public frmFlash As frmChat
	Public bUnitMoveShowReport As Boolean
	
	Public HavePlanes As Boolean
	Public HaveShips As Boolean
	Public HaveLands As Boolean
	Public HaveNukes As Boolean
	
	'State flags
	Public RecordingScriptFile As Boolean
	Public ServerQuery As Boolean
	Public CarrierDumpRequested As Boolean
	Dim DisplayCmd As Boolean
	Dim DatabaseUpdateinProgress As Boolean
	Dim BatchExploreinProgress As Boolean
	Dim BatchFileinProgress As Boolean
	Dim AutoMoveinProgress As Boolean
	Dim AntiSubPatrolinProgress As Boolean
	Dim AutoMoveLook As Boolean
	Dim AutoMoveView As Boolean
	Dim UpdateTimeStamp As Boolean
	Dim OutputToFile As Boolean
	Dim SubmitedfromCmdWindow As Boolean
	Dim BundleReports As Boolean
	Dim BatchFireAbort As Short
	Dim bNPHourlyActivity As Boolean '112903 rjk: Added for generating HourlyActivity Report from Newspaper
	Dim bNPNoReport As Boolean '120203 rjk: Allows the suppress of newspaper report
	Public bSetUpdateTime As Boolean 'set for when xdump updates 0 command is submitted
	Public bUpdateEnabled As Boolean 'set to reflect the value in xdump game
	
	Dim Password As String
	Public LoginUser As String
	Dim bKillSession As Boolean
	Dim bListSanc As Boolean
	Dim clstIndex As Short
	' Dim nFontSize As Integer   removed efj 8/2003
	Dim NewLineCount As Short
	Dim strReply1 As String
	Dim strReply2 As String
	Dim OutputFileNumber As Short
	Dim LogFileNumber As Short
	Dim LastFlash As String
	
	'Plane result variables
	Dim AirCombatResults As Boolean
	' Dim AirCombatOpponent As Boolean   removed efj 8/2003
	
	'Special bomb reply strings
	Dim bombreply1 As String
	Dim bombreply2 As String
	
	'Special attack parameters
	Dim Attack_Mil As Short
	Dim Attack_Occ As Short
	Dim Attack_Unit As String
	Dim Attack_Occ_Unit As String
	Dim Attack_Probe As String
	Dim Attack_DefMil As Short
	Dim Attack_Sector As String
	
	'Token for WinEmpireHub
	Public bEmpireHub As Boolean '150504 rjk: Add for showing chat window
	Enum enumToken
		tsIdle
		tsSent
		tsReceived
	End Enum
	Public tokenStatus As enumToken
	
	Public Sub ConnectToServer()
		Dim strCurr As String
		Dim strMRUAdd As String
		Dim n As Short
		
		'Save game in MRU
		strMRUAdd = GameName
		For n = 1 To 5
			strCurr = GetSetting(APPNAME, "MRUFiles", "MRU" & CStr(n), vbNullString)
			SaveSetting(APPNAME, "MRUFiles", "MRU" & CStr(n), strMRUAdd)
			strMRUAdd = strCurr
			If strMRUAdd = GameName Then
				Exit For
			End If
		Next n
		
		'Get the empire login info
		frmEmpireLogin.go() 'drk 5/25/03: added.  before, frmEmpireLogin depended on the Load event to do some
		'processing, but that didn't work so well if the form was already loaded but not shown
		
		frmEmpireLogin.ShowDialog()
		LoginIntoServer()
	End Sub
	
	Public Sub LoginIntoServer()
		Dim strBuffer As String
		
		If frmEmpireLogin.ConnectedtoHost Then
			ExpectSpecificResponse = READ_STATE_FIRSTSERVERINIT
			ConnectedtoHost = True
			CmdinProgress = True
			CountryName = frmEmpireLogin.txtUserName.Text
			Password = frmEmpireLogin.txtPassword.Text
			LoginUser = frmEmpireLogin.txtUser.Text
			If frmEmpireLogin.chkKill.CheckState = System.Windows.Forms.CheckState.Checked Then
				bKillSession = True
			Else
				bKillSession = False
			End If
			
			If frmEmpireLogin.chkSanc.CheckState = System.Windows.Forms.CheckState.Checked Then
				bListSanc = True
			Else
				bListSanc = False
			End If
			
			If frmEmpireLogin.chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked Then
				bAutoUpdate = True
			Else
				bAutoUpdate = False
			End If
			
			If frmEmpireLogin.chkAutoRead.CheckState = System.Windows.Forms.CheckState.Checked Then
				bAutoRead = True
			Else
				bAutoRead = False
			End If
			
			If frmEmpireLogin.chkDeity.CheckState = System.Windows.Forms.CheckState.Checked Then
				bDeity = True
			Else
				bDeity = False
			End If
			
			If frmEmpireLogin.chkUTF8.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				bUTF8 = True
			Else
				bUTF8 = False
			End If
			
			If frmEmpireLogin.chkHTTP.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				bXMLRPC = True
			Else
				bXMLRPC = False
			End If
			frmHTTPSettings.SaveXMLRPCParameters()
			
			'In case data has already been received, try to process
			ReadStringFromSocket(strBuffer)
			If Len(strBuffer) > 0 Then
				ProcessServerResponse(strBuffer)
			End If
		Else
			If Len(CountryName) = 0 Then
				CountryName = frmEmpireLogin.txtUserName.Text
			End If
			
			If frmEmpireLogin.chkDeity.CheckState = System.Windows.Forms.CheckState.Checked Then
				bDeity = True
			Else
				bDeity = False
			End If
			
			ConnectedtoHost = False
		End If
		
	End Sub
	Public Sub CheckLostConnection()
		Dim Reply As Short
		'Check for Connection
		If bRelogin And Not ConnectedtoHost And Not frmEmpireLogin.bOffline Then
			frmEmpireLogin.ConnectedtoHost = False
			frmEmpireLogin.Timer1.Interval = 350
			Exit Sub
		End If
		Do While Not ConnectedtoHost And Not frmEmpireLogin.bOffline '092603 rjk: Added Offline flag and add note to prompt
			Reply = MsgBox("Warning: Client has lost its connection to the empire server.  " & vbCrLf & "Use Ignore to work Offline", MsgBoxStyle.AbortRetryIgnore + MsgBoxStyle.Exclamation, "Connection Error")
			Select Case Reply
				Case MsgBoxResult.Abort
					End 'End the program at this point
				Case MsgBoxResult.Ignore
					'092603 rjk: Added Offline Mode
					frmDrawMap.Text = My.Application.Info.Title & " - " & GameName & " (Offline)"
					frmEmpireLogin.bOffline = True
					If Not frmDrawMap.MapValid Then '092703 rjk: For offline mode when the server is disable
						frmDrawMap.MapValid = True
						frmDrawMap.DrawHexes()
					End If
					Exit Do
				Case MsgBoxResult.Retry
					'            ConnectToServer
					frmEmpireLogin.ConnectedtoHost = False
					frmEmpireLogin.Timer1.Interval = 350
					frmEmpireLogin.Visible = True
					bRetry = True
					Exit Sub
					
				Case Else
					Exit Do 'Should never hit this, but this prevents inf. loop
					'   just in case it does
			End Select
		Loop 
	End Sub
	
	Private Sub frmEmpCmd_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		ConnectedtoHost = False
		BatchExploreinProgress = False
		ServerQuery = False
		DisplayCmd = True
		SubmittedFromCommandLine = False
		'Create Command Queue
		CmdQ = New EmpCmdQueue
		
		HavePlanes = True
		HaveShips = True
		HaveLands = True
		HaveNukes = True
		
		'Reset Attack Variables
		Attack_Mil = -1
		Attack_Occ = -1
		
		'Verify Directory Exists
		Dim strlog As String
		If VerifySubDirectory("Game Logs", True) Then
			strlog = My.Application.Info.DirectoryPath & "\Game Logs\" & GameName & " " & CStr(Year(Now)) & "-" & VB6.Format(Month(Now), "00") & "-" & VB6.Format(VB.Day(Now), "00") & ".log"
			
			'open the new log
			LogFileNumber = FreeFile
			'if it already exists, append to it
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Len(Dir(strlog)) > 0 Then
				FileOpen(LogFileNumber, strlog, OpenMode.Append)
			Else
				FileOpen(LogFileNumber, strlog, OpenMode.Output)
			End If
			PrintLine(LogFileNumber, "+++ Log opened at " & VB6.Format(Now, "hh:mm:ss") & " +++")
			PrintLine(LogFileNumber, vbNullString)
		End If
		FlashLogFileNumber = -1
		
	End Sub
	
	Private Sub frmEmpCmd_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		'Only unload if form Draw says OK
		'UPGRADE_ISSUE: Constant vbFormCode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		If UnloadMode = System.Windows.Forms.CloseReason.UserClosing Or UnloadMode = vbFormCode Then
			If Not frmDrawMap.ShutDown Then
				Cancel = 1
				Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
				Exit Sub
			End If
		End If
		ExecuteEmpireCmd("bye") 'sign off server
		
		'close log file
		FileClose(LogFileNumber)
		CloseFlashLog()
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event frmEmpCmd.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmEmpCmd_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		'When we resize, we must reset list and text box
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Exit Sub
		End If
		
		If VB6.PixelsToTwipsY(Me.Height) < VB6.PixelsToTwipsY(Text1.Height) * 6 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Text1.Height) * 6)
		End If
		
		'If Me.Width < Label1.Width * 2 Then
		'    Me.Width = Label1.Width * 2
		'End If
		
		List1.SetBounds(0, 0, Me.ClientRectangle.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 2 * VB6.PixelsToTwipsY(Text1.Height)))
		Text1.SetBounds(0, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(List1.Height) + 50), Me.ClientRectangle.Width, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
		imgLights.SetBounds(0, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Text1.Top) + VB6.PixelsToTwipsY(Text1.Height) + VB6.PixelsToTwipsY(imgLights.Height) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Label1.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(imgLights.Left) + VB6.PixelsToTwipsX(imgLights.Width)), imgLights.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
	
	Private Sub frmEmpCmd_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'If F2 is pressed, repeat last command
		If KeyCode = System.Windows.Forms.Keys.F2 Then
			SubmitLastCommand()
			SubmitedfromCmdWindow = True
		End If
		
		'If F9 is pressed, move back thru command list
		'102503 rjk: Added shift F9 moves forward thru the command list
		If KeyCode = System.Windows.Forms.Keys.F9 Then
			If Not ((Shift And VB6.ShiftConstants.ShiftMask) = VB6.ShiftConstants.ShiftMask) Then
				Text1.Text = GetPreviousCommand
			Else
				Text1.Text = GetNextCommand
			End If
		End If
		
		'check and see if the control key is down
		If Not ((Shift And VB6.ShiftConstants.CtrlMask) = 2) Then
			Exit Sub
		End If
		On Error Resume Next
		
		If KeyCode = System.Windows.Forms.Keys.D1 Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Normal
			Me.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D2 Then
			frmTelegram.Activate()
		End If
		If KeyCode = System.Windows.Forms.Keys.D3 Then
			frmTelegram.WindowState = System.Windows.Forms.FormWindowState.Normal
			frmTelegram.Activate()
		End If
		
	End Sub
	
	Public Function GetPreviousCommand() As String
		GetPreviousCommand = CmdQ.GetPreviousCommand(clstIndex)
		clstIndex = clstIndex - 1
		If clstIndex < 1 Then
			clstIndex = CmdQ.PreviousCommandCount
		End If
	End Function
	
	'102503 rjk: Added to support shift F9 - next command
	Public Function GetNextCommand() As String
		If CmdQ.PreviousCommandCount = 1 Then
			GetNextCommand = CmdQ.GetPreviousCommand(1)
			Exit Function
		End If
		
		If clstIndex = CmdQ.PreviousCommandCount Then
			GetNextCommand = CmdQ.GetPreviousCommand(2)
			clstIndex = 1
		ElseIf clstIndex = CmdQ.PreviousCommandCount - 1 Then 
			GetNextCommand = CmdQ.GetPreviousCommand(1)
			clstIndex = CmdQ.PreviousCommandCount
		Else
			GetNextCommand = CmdQ.GetPreviousCommand(clstIndex + 2)
			clstIndex = clstIndex + 1
		End If
	End Function
	
	Public Sub SubmitLastCommand()
		SubmitFromCommandLine(CmdQ.GetPreviousCommand(CmdQ.PreviousCommandCount))
	End Sub
	Public Sub SubmitFromCommandLine(ByRef strCmd As String)
		
		Dim nSector As Short
		Dim nShip As Short
		Dim nLand As Short
		Dim nNuke As Short
		Dim nPlane As Short
		Dim nLost As Short
		Dim nCountry As Short '100703 rjk: Added to detect the change command and submit a nation command to update country name
		Dim prefix As String
		
		SubmitEmpireCommand("cl1", False)
		SubmitEmpireCommand(strCmd, True)
		SubmitEmpireCommand("cl2 ", False)
		
		If Not bAutoUpdate Or Len(strCmd) < 3 Then
			Exit Sub
		End If
		
		'submit data updates based on the commands entered
		prefix = VB.Left(Trim(strCmd), 3)
		'100603 rjk: Removed duplicates from nSector list.
		nSector = InStr("ant arm ass att boa bom bui buy con del dem des dis dro enl exp " & "fir fue gri imp llo lmi loa lun mov par rec sel sho spy sta sto sup swe " & "ter thr unl upg wip wor ", prefix)
		
		nLand = InStr("bui fue llo lmi loa lun sup upg wor arm for lra lre " & "lte mar mor nav ten unl uns", prefix)
		
		nShip = InStr("ass boa bui fle fol fue loa lte min mqu nam nav ord ret sai ten tor unl upg", prefix)
		
		nPlane = InStr("arm bom bui dis dro fly har lau lte loa nav par rec sat swe ten tra unl upg win", prefix)
		
		nNuke = InStr("bui arm dis tra", prefix) '092303 rjk: Added tra for transporting nukes
		
		nLost = InStr("ant scr scu mar ass att nav", prefix)
		
		nCountry = InStr("cha", prefix)
		
		'fire is special.  You need to figure out what fired
		If prefix = "fir" Then
			nLand = InStr(4, strCmd, " l")
			nShip = InStr(4, strCmd, " sh")
			nSector = InStr(4, strCmd, " se")
			'if you can't figure it out, dump them all
			If nLand + nShip + nSector = 0 Then
				nLand = 1
				nShip = 1
				nSector = 1
			End If
		End If
		
		
		If (nSector + nLand + nShip + nPlane + nNuke + nCountry) > 0 Then
			Me.SubmitEmpireCommand("db1", False)
			If nSector > 0 Then
				GetSectors()
				'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
				GetCurrentStrength(tsSectors)
			End If
			
			If nPlane > 0 Then
				GetPlanes()
			End If
			
			If nLand > 0 Then
				GetLandUnits()
			End If
			
			If nShip > 0 Then
				GetShips()
				GetOContent() '110605 rjk: Added ability to get OContent for Sea Sectors
			End If
			
			If nNuke > 0 Then
				GetNukes()
			End If
			
			If nLost > 0 Then
				GetLost()
			End If
			
			If nCountry > 0 Then
				GetNation()
			End If
			
			Me.SubmitEmpireCommand("db2", False)
		End If
	End Sub
	
	Private Sub frmEmpCmd_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		' Dim n As Integer   removed efj 8/2003
		' When Enter is pressed, process everything
		If KeyAscii = 13 Then
			If ServerQuery Then
				ExecuteEmpireCmd((Text1.Text))
				SubmitedfromCmdWindow = True
			Else
				'now submit the command
				SubmitFromCommandLine((Text1.Text))
				SubmitedfromCmdWindow = True
			End If
			Text1.Text = vbNullString
			KeyAscii = 0
		End If
		
		'Trap Control C and Copy to Clipboard
		If KeyAscii = 3 Then
			CopyListBoxToClipboard(List1)
		End If
		
		'if an escape key is hit when a query is up
		If KeyAscii = 27 And ServerQuery Then
			ExecuteEmpireCmd("ctld")
			VB6.SetItemString(List1, List1.Items.Count - 1, VB6.GetItemString(List1, List1.Items.Count - 1) & " ctld")
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Public Sub SubmitEmpireCommand(ByRef strCmd As String, Optional ByRef bAddtoList As Object = Nothing)
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(bAddtoList) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object bAddtoList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bAddtoList = True
		End If
		
		'Reset Submitted Source State Flag
		SubmitedfromCmdWindow = False
		
		If strCmd = vbNullString Then
			Exit Sub
		End If
		
		'Put into command list
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object bAddtoList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If bAddtoList Then
			CmdQ.AddtoHistory(strCmd)
			clstIndex = CmdQ.PreviousCommandCount
		End If
		
		CmdQ.AddCommand(strCmd)
		
		'If there are already other pending commands
		'or if there is a command in progress, leave
		If CmdQ.CmdsPending > 1 Or CmdinProgress Then
			Exit Sub
		End If
		
		'Execute the command
		TransferEmpireCmd()
	End Sub
	
	
	Public Sub ExecuteEmpireCmd(ByRef strCmd As String)
		On Error GoTo Empire_Error
		
		Dim iCapX As Short
		Dim iCapY As Short
		Dim n As Short
		Dim strID As String
		
		'remove Spaces
		strCmd = Trim(strCmd)
		
		'Ignore empty strings
		If Len(strCmd) = 0 And Not ServerQuery Then
			Exit Sub
		End If
		
		'Ignore comments, responses in batch files
		If VB.Left(strCmd, 2) = "//" Or VB.Left(strCmd, 1) = "&" Or VB.Left(strCmd, 1) = "'" Then
			Exit Sub
		End If
		
		'Reset time idle
		TimeIdle = 0
		
		'check for specially trapped commands
		strID = VB.Left(strCmd, 3)
		
		'if this is a script file being recorded - don't process command
		Dim strSkip As String
		If RecordingScriptFile Then
			'you want to skip the various dumps in a script file
			strSkip = "dum pdu ldu sdu ndu db1 db2 bfa bf1 bf2 bma map upd xdu"
			If InStr(strSkip, strID) = 0 Then
				frmScript.AddCommand(strCmd)
			End If
			Exit Sub
		End If
		
		Dim n2 As Short
		Dim strn As String
		Dim stru As String
		Dim nr As Short
		Dim no As Short
		Dim sc As Short
		Dim numberToEachSector As Short
		Dim exploreMode As WinAceRoutines.enumExploreMode
		Dim strTemp As String '090103 rjk: Added for recursion call
		Dim FileName, Line As String
		Dim fno As Short
		If Not ServerQuery Then
			Select Case strID
				Case "dum" 'Sector Dump
					If ForceUpdates Or bAutoUpdate Then
						ExpectSpecificResponse = READ_STATE_SECTOR_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
				Case "pdu" 'Plane Dump
					If HavePlanes And (ForceUpdates Or bAutoUpdate) Then
						ExpectSpecificResponse = READ_STATE_PLANE_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
				Case "ldu" 'Land Dump
					If HaveLands And (ForceUpdates Or bAutoUpdate) Then
						ExpectSpecificResponse = READ_STATE_LAND_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
					
				Case "sdu" 'Ship Dump
					If HaveShips And (ForceUpdates Or bAutoUpdate) Then
						ExpectSpecificResponse = READ_STATE_SHIP_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
				Case "ndu" 'Nuke Dump
					If HaveNukes And (ForceUpdates Or bAutoUpdate) Then
						ExpectSpecificResponse = READ_STATE_NUKE_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
				Case "los" 'Lost Dump
					If ForceUpdates Or bAutoUpdate Then
						ExpectSpecificResponse = READ_STATE_LOST_DUMP
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					Else
						NextCommand()
						Exit Sub
					End If
				Case "bma" 'Bmap Dump
					'if we are suppressing the bmap, exit
					If frmOptions.bSuppressBmapRefresh Then
						NextCommand()
						Exit Sub
					End If
					
					If DatabaseUpdateinProgress Then
						ExpectSpecificResponse = READ_STATE_FULL_BMAP
						DisplayCmd = False
					Else
						SetReport("B-map")
					End If
				Case "map" 'Map Dump
					If DatabaseUpdateinProgress Then
						ExpectSpecificResponse = READ_STATE_FULL_BMAP
						DisplayCmd = False
					Else
						SetReport("Map")
					End If
				Case "nma", "sma", "lma", "pma", "sbm", "lbm", "pbm" 'Map Dump
					SetReport(strCmd)
					
				Case "xdu"
					ExpectSpecificResponse = READ_STATE_XDUMP
					If DatabaseUpdateinProgress Then
						DisplayCmd = False
						If InStr(strCmd, "*") > 0 Then
							UpdateTimeStamp = True
						End If
					End If
					If strCmd = "xdump updates 0" Then
						bSetUpdateTime = True
					End If
					
				Case "acc" 'Accept Report
					SetReport("Acceptance Report")
					
				Case "att", "ass" 'Attack, Assult Report
					ExpectSpecificResponse = READ_STATE_ATTACK
					'SetReport "Attack Report"
					
				Case "bom", "fly", "dro" 'Bombing  Report
					If Not SubmittedFromCommandLine Then
						frmReport.AddReportLine(strCmd)
						ExpectSpecificResponse = READ_STATE_BOMB
						DisplayCmd = False
					Else
						ExpectSpecificResponse = READ_STATE_BOMBCL
					End If
					
				Case "bre" 'break sanctuary
					ExpectSpecificResponse = READ_STATE_BREAK
					
				Case "bud" 'Budget Report
					SetReport("Budget Report")
					
				Case "cen" 'Census Report
					SetReport("Census Report")
					
				Case "coa" 'Coastwatch Report
					SetReport("Coastwatch Report")
					ExpectSpecificResponse = READ_STATE_COASTWATCH
					
				Case "cou" 'Country Report
					SetReport("Country Report")
					
				Case "com" 'Commodity Report
					SetReport("Commodity Report")
					
				Case "cut" 'Cutoff Report 092303 rjk: Added to capture cutoff report
					SetReport("Cutoff Report")
					
				Case "edi" 'edit
					ExpectSpecificResponse = READ_STATE_EDIT
					'check for reply string
					n = InStr(strCmd, ":")
					If n > 0 Then
						strReply1 = Mid(strCmd, n + 1)
						strCmd = VB.Left(strCmd, n - 1)
					End If
					
				Case "exp" 'Explore
					If BatchExploreinProgress Then
						ExpectSpecificResponse = READ_STATE_EXPLORE_CIRCLE
						DisplayCmd = True
					End If
					
				Case "fin" 'Financial Report
					SetReport("Financial Report")
					
				Case "fir", "tor" 'fire, torpedo
					If BatchFileinProgress And BatchFireAbort = BF_ABORTREMAINING Then
						NextCommand()
						Exit Sub
					Else
						ExpectSpecificResponse = READ_STATE_FIRE_SEQ
					End If
					
				Case "hid", "pee" 'hidden (deity command)
					SetReport("Hidden Attributes")
					
				Case "inf" 'Info Report
					SetReport("Empire Help")
					
				Case "lan" 'Land
					SetReport("Land Units")
					
				Case "lau"
					If Not SubmittedFromCommandLine Then
						ExpectSpecificResponse = READ_STATE_LAUNCH
						DisplayCmd = False
					End If
					
				Case "lca" 'Land Cargo
					SetReport("Land Cargo")
					
				Case "led" 'Loan Ledger
					SetReport("Loan Ledger")
					
				Case "lev" 'Level Report 092303 rjk: Added to capture level report
					SetReport("Level Report")
					
				Case "lis" 'List of Commands
					SetReport("List of Commands")
					
				Case "llo" 'Land Look/Load
					If VB.Left(strCmd, 4) = "lloo" Then
						SetReport("Land Lookout Reports")
						ExpectSpecificResponse = READ_STATE_LOOK
					End If
					
				Case "loo" 'Look
					SetReport("Ship Lookout Reports")
					ExpectSpecificResponse = READ_STATE_LOOK
					
				Case "lst" 'Land Stat
					SetReport("Land Statistics")
					
				Case "mar" 'Market Report or march
					If VB.Left(strCmd, 4) = "mark" Then
						SetReport("Market Report")
					Else
						If Not SubmittedFromCommandLine Then
							ExpectSpecificResponse = READ_STATE_UNIT_MOVE
							bUnitMoveShowReport = False
						End If
					End If
					
				Case "mis" 'Mission
					n = InStr(Trim(strCmd), " ")
					If n > 0 Then
						Select Case LCase(VB.Left(Trim(Mid(strCmd, n)), 1))
							Case "l"
								ExpectSpecificResponse = READ_STATE_LANDMISSION
							Case "s"
								ExpectSpecificResponse = READ_STATE_SHIPMISSION
							Case "p"
								ExpectSpecificResponse = READ_STATE_PLANEMISSION
						End Select
					End If
					
				Case "mot" 'Message of the Day
					SetReport("Message of the Day")
					
				Case "nat" 'Nation Report
					If Not DatabaseUpdateinProgress Then
						SetReport("Nation Report")
					End If
					If Not SubmittedFromCommandLine Then
						ExpectSpecificResponse = READ_STATE_NATION
						DisplayCmd = False '100103 rjk: Added so when enlist or demobilize commands are not bumped by nation update (mil. reserves)
					End If
					
				Case "nav" 'navigate
					If Not SubmittedFromCommandLine Then
						ExpectSpecificResponse = READ_STATE_UNIT_MOVE
						bUnitMoveShowReport = False
						If DatabaseUpdateinProgress Then
							DisplayCmd = False
						End If
					End If
					
				Case "new"
					'newspaper or newefficiency
					If VB.Left(strCmd, 4) = "news" Then
						SetReport("Empire Newspaper")
						ExpectSpecificResponse = READ_STATE_NEWSPAPER '112903 rjk: Added for intelligence gathering.
					ElseIf VB.Left(strCmd, 4) = "newe" Then  '101803 rjk: Added Neweff report
						SetReport("New Efficiency Report")
					End If
					
				Case "nuk" 'Nukes
					SetReport("Nuclear Devices")
					
				Case "ord" 'order
					If BatchFileinProgress Then
						ExpectSpecificResponse = READ_STATE_ORDER
					End If
					
				Case "ori" 'origin
					ExpectSpecificResponse = READ_STATE_ORIGIN
					
				Case "par" 'partroop,
					If BatchFileinProgress Then
						ExpectSpecificResponse = READ_STATE_FIRE_SEQ
					Else
						SetReport("Mission Results")
					End If
					
				Case "pca" 'Plane Cargo
					SetReport("Plane Cargo")
					
				Case "pow" 'Power Report
					SetReport("Power Report")
					
				Case "pla" 'Planes
					If VB.Left(strCmd, 4) = "play" Then
						If DatabaseUpdateinProgress Then
							DisplayCmd = False
						Else
							SetReport("Player Report")
						End If
						ExpectSpecificResponse = READ_STATE_PLAYERS '120103 rjk: Added for displaying current players
					Else
						SetReport("Plane")
					End If
					
				Case "pro" 'Production Report
					SetReport("Production Report")
					
				Case "pst" 'Plane Stat
					SetReport("Plane Statistics")
					
				Case "qor" 'query orders
					SetReport("Ship Orders")
					ExpectSpecificResponse = READ_STATE_SHOWORDER
					
				Case "rad" 'Radar Report
					SetReport("Radar Report")
					
				Case "rea"
					'Seperate read vs realm
					If VB.Left(strCmd, 4) = "read" Then
						ExpectSpecificResponse = READ_STATE_TELEGRAM_READ_CL
						strLastProductionReport = vbNullString
						If Not SubmittedFromCommandLine Then
							ExpectSpecificResponse = READ_STATE_TELEGRAM_READ
							strReply1 = VB.Right(Trim(strCmd), 1)
							If Not (strReply1 = "y" Or strReply1 = "n") Then
								strReply1 = vbNullString
							End If
							DisplayCmd = False
						End If
					End If
					
				Case "rec", "swe" ' recon, sweep
					SetReport("Recon Report")
					ExpectSpecificResponse = READ_STATE_RECON
					AntiSubPatrolinProgress = False
					
				Case "rel" 'Relations Report
					'110303 rjk: Change processing of relations to always fill Nations object information not just from database update.
					'111603 rjk: Removed report
					'112503 rjk: Added back report
					If Not DatabaseUpdateinProgress Then
						SetReport("Relations Report")
					End If
					ExpectSpecificResponse = READ_STATE_REQ_RELATIONS
					If Not SubmittedFromCommandLine Then
						DisplayCmd = False
					End If
					
				Case "rep" 'Technology/Education/Happiness Report - also "repay"
					If VB.Left(strCmd, 4) = "repo" Then
						'100703 rjk: READ_STATE_REQ_REPORT not required a nation list is done in relations now.
						'If DatabaseUpdateinProgress Then
						'    ExpectSpecificResponse = READ_STATE_REQ_REPORT
						'    DisplayCmd = False
						'Else
						SetReport("Tech Report")
						'End If
					End If
					
				Case "res" 'Resource Report
					If VB.Left(strCmd, 4) = "reso" Then
						SetReport("Resource Report")
					End If
					
				Case "rou"
					SetReport("Route Report")
					
				Case "sat" 'Country Report
					SetReport("Satellite Report")
					ExpectSpecificResponse = READ_STATE_RECON
					AntiSubPatrolinProgress = False
					
				Case "car" 'Ship Cargo 092303 rjk: switched to "car" from "sca"
					SetReport("Ship Cargo")
					
				Case "scu", "scr" '040404 rjk: Added for scuttle and scrap
					If Not SubmittedFromCommandLine Then
						ExpectSpecificResponse = READ_STATE_SCUTTLE
						DisplayCmd = False
					End If
					
				Case "shi" 'Ship
					SetReport("Ships")
					
				Case "sst" 'Ship Stat
					SetReport("Ship Statistics")
					
				Case "sho" 'Show Report
					If VB.Left(strCmd, 4) = "show" Then
						If VB.Left(strCmd, 8) = "show upd" Then
							SetReport("Update Schedule")
						Else
							If DatabaseUpdateinProgress Then '101403 rjk: Added to suppress show sect s during startup
								DisplayCmd = False
							Else
								SetReport(strCmd)
							End If
							ExpectSpecificResponse = READ_STATE_SHOW
						End If
					End If
					
				Case "sky" 'Sky watch Report
					SetReport("Skywatch Report")
					
				Case "sor" 'show orders
					SetReport("Ship Orders")
					ExpectSpecificResponse = READ_STATE_SHOWORDER
					
				Case "son" 'Sonar
					SetReport("Sonar Report")
					ExpectSpecificResponse = READ_STATE_SONAR
					
				Case "spy" 'Spy 101803 rjk: Combined duplicate cases
					SetReport("Spy Report")
					ExpectSpecificResponse = READ_STATE_SPY
					
				Case "sta" 'Starvation or Start
					If VB.Left(strCmd, 5) = "starv" Then
						If Not SubmittedFromCommandLine Then
							SetReport("Starvation Report")
							DisplayCmd = False
							ExpectSpecificResponse = READ_STATE_STARVATION
						End If
					Else
						'Start Command
					End If
					
				Case "str" '101703 rjk: Strength Report added and processing for Sector db and Sector display
					If Not DatabaseUpdateinProgress Then
						SetReport("Strength Report")
					End If
					If Not SubmittedFromCommandLine Then
						DisplayCmd = False
					End If
					ExpectSpecificResponse = READ_STATE_STRENGTH
					
				Case "tel", "ann", "fla", "tur" 'telegram, announcment
					If Not (strID = "tur" And strCmd <> "turn mess") Then
						If Not SubmittedFromCommandLine Then
							frmReport.AddReportLine(strCmd)
							If strID = "fla" Then
								ExpectSpecificResponse = READ_STATE_FLASH_WRITE
							Else
								ExpectSpecificResponse = READ_STATE_TELEGRAM_WRITE
							End If
							DisplayCmd = False
						Else
							If strID = "fla" Then
								ExpectSpecificResponse = READ_STATE_FLASH_WRITE
							End If
						End If
					End If
					
				Case "tes" 'test
					SetReport("Test Movement Report")
					
					'101703 rjk: Removed trade report, as trade is interactive and does not work as report
					'Case "tra"  'Trade/Transport Report
					'    If Left$(strCmd, 4) = "trad" Then
					'        SetReport "Trade Report"
					'    Else
					'        'transport Command
					'    End If
					'
				Case "tre" 'Treaty 100303 rjk: Added to produce the output in report form
					SetReport("Treaty")
					
				Case "upd" 'Update Dump ''drk: fixme here
					If Not SubmittedFromCommandLine Then
						ExpectSpecificResponse = READ_STATE_NEXT_UPDATE
						DisplayCmd = Not DatabaseUpdateinProgress
					End If
					
				Case "ver" 'Version Report
					SetReport("Game Version")
					ExpectSpecificResponse = READ_STATE_VERSION
					
				Case "wir"
					SetReport("Wire Report")
					frmDrawMap.sb1.Items.Item("Anno").Text = " "
					ExpectSpecificResponse = READ_STATE_WIRE
					
					'added 7/20/02 drk
				Case "exe"
					
					n = InStr(strCmd, " ")
					If (n > 0) Then
						FileName = VB.Right(strCmd, Len(strCmd) - n)
						fno = FreeFile
						
						On Error GoTo notfound
						
						FileOpen(fno, FileName, OpenMode.Input)
						While (Not EOF(fno))
							Line = LineInput(fno)
							Me.SubmitEmpireCommand(Line, False)
						End While
						FileClose((fno))
					End If
					
					NextCommand()
					Exit Sub
					
notfound: 
					MsgBox("File Error.  File not found or error reading file!", MsgBoxStyle.OKOnly)
					NextCommand()
					Exit Sub
					
					'Dummy Codes for setting data update variables - not passed on to Empire Server
				Case "dbr" '(Reset data base)
					ResetDataBase()
					NextCommand()
					Exit Sub
				Case "db1"
					DatabaseUpdateinProgress = True
					NextCommand()
					Exit Sub
				Case "db2"
					DatabaseUpdateinProgress = False
					ForceUpdates = False
					If Not frmDrawMap.MapValid Then
						If frmOptions.bCapital And ParseSectors(iCapX, iCapY, CapSect) Then
							frmDrawMap.MoveMap(iCapX, iCapY)
						Else
							frmDrawMap.MoveMap(0, 0)
						End If
						frmDrawMap.Cursor = System.Windows.Forms.Cursors.Default
					Else
						frmDrawMap.DrawHexes()
					End If
					
					'if you want to know when a database update has finished, add your callbacks here
					If frmRelationsGrid.Visible Then frmRelationsGrid.callback()
					
					NextCommand()
					Exit Sub
					
					'Dummy Codes for setting up exploration circle
				Case "ex1"
					BatchExploreinProgress = True
					BatchFileinProgress = True
					NextCommand()
					Exit Sub
					
				Case "ex2", "ex4"
					BatchExploreinProgress = False
					BatchFileinProgress = False
					If strID = "ex2" Then
						frmDrawMap.lstCmdResult.Items.Add("Updating Map....")
					Else
						frmDrawMap.lstCmdResult.Items.Add("Exploration Complete !")
					End If
					'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
					NextCommand()
					Exit Sub
					
				Case "ex3" 'nova command - keeps the exploration going
					'090103 rjk: added exploreMode to the ExploreOut function
					
					n = InStr(strCmd, ":")
					strn = Trim(Mid(strCmd, 5, n - 5))
					n2 = InStr(n + 1, strCmd, ":")
					stru = Trim(Mid(strCmd, n + 1, n2 - n - 1))
					n = InStr(n2 + 1, strCmd, ":")
					nr = CShort(Trim(Mid(strCmd, n2 + 1, n - n2 - 1)))
					n2 = InStr(n + 1, strCmd, ":")
					no = CShort(Trim(Mid(strCmd, n + 1, n2 - n - 1)))
					n = InStr(n2 + 1, strCmd, ":")
					sc = CShort(Trim(Mid(strCmd, n2 + 1, n - n2 - 1)))
					n2 = InStr(n + 1, strCmd, ":")
					numberToEachSector = CShort(Trim(Mid(strCmd, n + 1, n2 - n - 1)))
					n = InStr(n2 + 1, strCmd, ":")
					exploreMode = CInt(Trim(Mid(strCmd, n2 + 1, n - n2 - 1)))
					CmdinProgress = True
					ExploreOut(strn, stru, nr, no, sc, numberToEachSector, exploreMode)
					CmdinProgress = False
					NextCommand()
					Exit Sub
					
					'Dummy Codes for setting up auto stop on nav, march
				Case "as1"
					AutoMoveinProgress = True
					AutoMoveLook = False
					AutoMoveView = False
					If Mid(strCmd, 4, 1) = "l" Then
						AutoMoveLook = True
					End If
					If Mid(strCmd, 4, 1) = "v" Or Mid(strCmd, 5, 1) = "v" Then
						AutoMoveView = True
					End If
					NextCommand()
					Exit Sub
					
					'Dummy Codes for launch
				Case "la1"
					strReply1 = Trim(Mid(strCmd, 5))
					NextCommand()
					Exit Sub
					
					'Dummy Codes for batch flying
				Case "ba1"
					SetReport("Flight Report")
					strReply1 = Trim(Mid(strCmd, 5, 3))
					strReply2 = Trim(Mid(strCmd, 8))
					bombreply1 = strReply1
					bombreply2 = strReply2
					If frmOptions.bClearResultNewCommand Then
						frmDrawMap.lstCmdResult.Items.Clear()
						frmDrawMap.lstCmdResult.Items.Add("Flying....")
					End If
					NextCommand()
					Exit Sub
				Case "ba2"
					strReply1 = vbNullString
					strReply2 = vbNullString
					bombreply1 = vbNullString
					bombreply2 = vbNullString
					BatchFileinProgress = False
					frmDrawMap.lstCmdResult.Items.Add("Mission Complete !")
					'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
					NextCommand()
					Exit Sub
					
					'Telegram batch
				Case "te1"
					strReply1 = Mid(strCmd, 5)
					strReply2 = "."
					NextCommand()
					Exit Sub
					
					'Dummy Codes for batch files
				Case "bf1", "bfa"
					BatchFileinProgress = True
					If Mid(strCmd, 3, 1) = "a" Then
						BatchFireAbort = BF_ABORTONRETURNFIRE
					Else
						BatchFireAbort = BF_NEVERABORT
					End If
					
					If frmOptions.bClearResultNewCommand Then
						frmDrawMap.lstCmdResult.Items.Clear()
						frmDrawMap.lstCmdResult.Items.Add("Executing Batch File")
					End If
					NextCommand()
					Exit Sub
				Case "bf2"
					BatchFileinProgress = False
					BatchFireAbort = BF_NEVERABORT
					frmDrawMap.lstCmdResult.Items.Add("Batch File Complete !")
					'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
					NextCommand()
					Exit Sub
					
					'Dummy Codes for setting Command Line indicator
				Case "cl1"
					SubmittedFromCommandLine = True
					NextCommand()
					Exit Sub
					
				Case "cl2"
					
					SubmittedFromCommandLine = False
					DisplayCmd = True
					If OutputToFile Then
						FileClose(OutputFileNumber)
						OutputToFile = False
					End If
					NextCommand()
					Exit Sub
					
				Case "cs1"
					DeleteAllRecords(rsSectors)
					NextCommand()
					Exit Sub
					
				Case "cu1"
					ClearUnitsFiles()
					NextCommand()
					Exit Sub
					
					'code to set attack header
				Case "at1"
					
					'Attack Sector
					strTemp = StringBetween(strCmd, "/sec=", ";")
					If Len(strTemp) > 0 Then
						If Trim(strTemp) <> Attack_Sector Then
							Attack_DefMil = -1
						End If
						Attack_Sector = Trim(strTemp)
					Else
						Attack_Sector = vbNullString
					End If
					
					'Attack Force, if Specified
					strTemp = StringBetween(strCmd, "/mil=", ";")
					If Len(strTemp) > 0 Then
						If strTemp = "ns" Then
							Attack_Mil = -2
						Else
							Attack_Mil = CShort(Trim(strTemp))
						End If
					Else
						Attack_Mil = -1
					End If
					
					'Occupying Force - if specified
					strTemp = StringBetween(strCmd, "/occ=", ";")
					If Len(strTemp) > 0 Then
						Attack_Occ = CShort(Trim(strTemp))
					Else
						Attack_Occ = -1
					End If
					
					'Use units in Attack
					strTemp = StringBetween(strCmd, "/unt=", ";")
					If Len(strTemp) > 0 Then
						Attack_Unit = Trim(strTemp)
					Else
						Attack_Unit = vbNullString
					End If
					
					'Occupy with Units
					strTemp = StringBetween(strCmd, "/ocu=", ";")
					If Len(strTemp) > 0 Then
						Attack_Occ_Unit = Trim(strTemp)
					Else
						Attack_Occ_Unit = vbNullString
					End If
					
					'Probe only
					n = InStr(strCmd, "/probe=")
					strTemp = StringBetween(strCmd, "/probe=", ";")
					If Len(strTemp) > 0 Then
						Attack_Probe = Trim(strTemp)
					End If
					
					NextCommand()
					Exit Sub
					
				Case "at2"
					'Reset Attack Variables
					Attack_Mil = -1
					Attack_Occ = -1
					Attack_Unit = vbNullString
					Attack_Probe = vbNullString
					Attack_Occ_Unit = vbNullString
					NextCommand()
					Exit Sub
					
				Case "br1"
					'Bundle reports
					BundleReports = True
					NextCommand()
					Exit Sub
					
				Case "br2"
					'Bundle reports
					BundleReports = False
					NextCommand()
					Exit Sub
					
				Case "np1" '112903 rjk: Added for HourlyActivity Report
					If InStr(strCmd, "Hourly") > 0 Then
						bNPHourlyActivity = True
						bNPNoReport = True
					ElseIf InStr(strCmd, "None") > 0 Then 
						bNPNoReport = True
					End If
					NextCommand()
					Exit Sub
				Case "np2" '112903 rjk: Added for HourlyActivity Report
					bNPHourlyActivity = False
					bNPNoReport = False
					NextCommand()
					Exit Sub
					
				Case "sc1" '040404 rjk: Added for automatic answer to scuttle
					strReply1 = "y"
					NextCommand()
					Exit Sub
			End Select
			
			'if its in a batch file - it goes to the command line
			If BatchFileinProgress Then
				DisplayCmd = True
				frmDrawMap.lstCmdResult.Items.Add("Command: " & strCmd)
				
				'Add to list box and reposition
			ElseIf DisplayCmd Then 
				If frmOptions.bClearResultNewCommand Then
					frmDrawMap.lstCmdResult.Items.Clear()
				End If
				frmDrawMap.lstCmdResult.Items.Add("Command: " & strCmd)
			End If
			EchoLine("Command: " & strCmd)
		Else
			VB6.SetItemString(List1, List1.Items.Count - 1, VB6.GetItemString(List1, List1.Items.Count - 1) & " " & strCmd)
		End If
		
		'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		List1.TopIndex = List1.NewIndex
		
		'Send the data to the Winsock Control
		SendStringtoServer(strCmd)
		If bEmpireHub Then
			Select Case tokenStatus
				Case enumToken.tsReceived
					tokenStatus = enumToken.tsIdle
				Case enumToken.tsIdle
					If Not CmdinProgress Then
						Debug.Print("ExecuteEmpireCmd: invalid tokenStatus")
					End If
				Case Else
					Debug.Print("ExecuteEmpireCmd: invalid tokenStatus")
					tokenStatus = enumToken.tsIdle
			End Select
		End If
		Exit Sub
Empire_Error: 
		EmpireError("ExecuteEmpireCmd", vbNullString, strCmd)
		
	End Sub
	Public Sub SendStringtoServer(ByRef strCmd As String)
		
		'Just in case connection breaks after the
		'ConnectedtoHost but before the SendData - rare but has
		'happened
		On Error GoTo SendStringtoServer_Exit
		
		'Send the data to the Winsock Control
		If ConnectedtoHost Then
			WriteStringToSocket(strCmd & vbLf)
			If ExpectSpecificResponse = READ_STATE_FLASH_WRITE Then
				FlashLog("(" & Me.LoginUser & ")" & strCmd)
			End If
			ServerQuery = False
			CmdinProgress = True
			UpdateLights()
		End If
		
SendStringtoServer_Exit: 
	End Sub
	
	Public Sub WriteStringToSocket(ByRef strCmd As String)
		Dim baUTF8() As Byte
		If bXMLRPC Then
			If Not xmlrpcServer Is Nothing Then
				xmlrpcServer.Send(strCmd)
			End If
		Else
			If bUTF8 Then
				baUTF8 = UTF8_Encode(strCmd)
				Winsock.SendData(baUTF8)
			Else
				Winsock.SendData(strCmd)
			End If
		End If
	End Sub
	
	Public Sub ReadStringFromSocket(ByRef strCmd As String)
		Dim baUTF8() As Byte
		If bXMLRPC Then
			If Not xmlrpcServer Is Nothing Then
				strCmd = xmlrpcServer.Receive
			End If
		Else
			If bUTF8 Then
				Winsock.GetData(baUTF8, VariantType.Array + VariantType.Byte)
				strCmd = UTF8_Decode(baUTF8)
			Else
				Winsock.GetData(strCmd, VariantType.String)
			End If
		End If
	End Sub
	Private Sub SetReport(ByRef ReportName As String)
		On Error GoTo Empire_Error
		
		If SubmittedFromCommandLine Then
			Exit Sub
		End If
		
		If Not BatchFileinProgress Then
			ExpectSpecificResponse = READ_STATE_REPORT
			DisplayCmd = False
			frmReport.Text = ReportName
			If Not BundleReports Then
				frmReport.ClearReport()
			End If
		End If
		
		If frmOptions.bClearResultNewCommand Then
			frmDrawMap.lstCmdResult.Items.Clear()
			frmDrawMap.lstCmdResult.Items.Add("Getting " & frmReport.Text)
		End If
		Exit Sub
Empire_Error: 
		EmpireError("SetReport", vbNullString, ReportName)
	End Sub
	
	Private Sub tmrXMLRPC_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrXMLRPC.Tick
		xmlrpcServer.HasData()
	End Sub
	
	Private Sub Winsock_DataArrival(ByVal eventSender As System.Object, ByVal eventArgs As AxMSWinsockLib.DMSWinsockControlEvents_DataArrivalEvent) Handles Winsock.DataArrival
		Dim strBuffer As String
		
		If frmEmpireLogin.bOffline Then '092703 rjk: Added for Offline mode, infinite loop otherwise
			Winsock.Close()
		End If
		
		If Not ConnectedtoHost Then
			Exit Sub
		End If
		
		'Retrieve the data
		ReadStringFromSocket(strBuffer)
		ProcessServerResponse(strBuffer)
	End Sub
	
	'I split the 2.1.0 ProcessServerResponse up into several smaller subs.
	'Any sub that is 1800 lines long is just too bloody long.
	'still too long though (~1600 lines). drk 7/03/02
	Public Sub ProcessServerResponse(ByRef strBuffer As String)
		Dim strData As String
		Dim strResponse As String
		Dim strCode As String
		Dim nCode As Short
		Dim n As Short
		Dim n2 As Short
		Dim strTemp As String
		Dim sx As Short
		Dim sy As Short
		Dim VARsec As Object
		Static strHoldOver As String
		Static Line As Short
		Dim shelldamage As Short
		Dim shellsector As String
		Static bXDumpServer As Boolean
		
		'On Error GoTo Empire_Error
		
		'Some times the fetch data request may fetch the server buffer before
		'a complete line is put in.  In these cases, we need to retain the extra
		'and merge it to the front of the next data request.
		
		strBuffer = strHoldOver & strBuffer
		strHoldOver = vbNullString
		
		'Process the data
		Dim mailcount As Short
		Dim strFile As String
		Dim bAppend As Boolean
		Static BmapFirstTimeFlag As Short
		Static bmaphead1 As String
		Static bmaphead2 As String
		Static bmaphead3 As String
		Static StartX As Short
		Static linelen As Short
		Dim temp As Short
		Dim Nation As Short
		Dim NationName As String
		Dim NationTowards As Short
		Static NationFrom As Short
		Dim NationFromName As String
		Dim NationRel As Short
		Dim n1 As String
		Static Looking As Boolean
		Static TimeUpdate As String
		Static TimeCurrent As String
		Dim totmin As Integer
		Dim sec As Integer
		Dim Min As Integer
		Static divider As Short
		Static UnitType As String
		Static ReportType As String
		Static VersionFirstTimeFlag As Short
		Dim strTempLine As String
		Dim HourIndex As Short
		Dim NationIndex As Short
		'UPGRADE_NOTE: hour was upgraded to hour_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim hour_Renamed As Integer
		While Len(strBuffer) > 0
			
			'Buffer could contain several lines, so seperate at line feeds and
			'process one at a time
			n = InStr(strBuffer, vbLf)
			If n = 0 Then
				'No line feed found means we do not have a complete line.  Leave
				'whats left in the Holdover string
				strData = vbNullString
				strHoldOver = strBuffer
				strBuffer = vbNullString
				Exit Sub 'Leave - more data is coming
			Else
				'Seperate one line for processing, drop the line feed, and leave whats left
				'in the buffer.
				strData = VB.Left(strBuffer, n - 1)
				If Len(strBuffer) = n Then
					strBuffer = vbNullString
				Else
					strBuffer = Mid(strBuffer, n + 1)
				End If
			End If
			
			'Get the text portion without the control characters
			If Len(strData) > 2 Then
				strResponse = VB.Right(strData, Len(strData) - 2) ' Make a copy of the string so we have one for sect command
			Else
				strResponse = vbNullString 'Error - we should never be here
			End If
			
			If Len(strData) > 0 Then
				'First Character is response code
				strCode = VB.Left(strData, 1)
				If strCode >= "a" And strCode <= "z" Then
					nCode = 10 + Asc(strCode) - Asc("a")
				ElseIf strCode >= "0" And strCode <= "9" Then 
					nCode = Asc(strCode) - Asc("0")
				Else
					MsgBox("Invalid Server Code - " & CStr(Asc(strCode)))
				End If
			Else
				nCode = EMP_C_DATA 'Error - we should never be here
			End If
			
			'Check for air combat records
			If Not AirCombatResults Then
				If InStr(strResponse, "strength int odds") > 0 Then
					AirCombatResults = True
					ParseAirCombat(strResponse)
				End If
			Else
				If InStr(Mid(strResponse, 50, 15), "abort") > 0 Or InStr(Mid(strResponse, 50, 15), "clear") > 0 Or InStr(Mid(strResponse, 50, 15), "shot") > 0 Then
					ParseAirCombat(strResponse)
				Else
					AirCombatResults = False
				End If
			End If
			
			'Check for Specific Messages
			If InStr(strResponse, "You lost your capital") > 0 Then
				EchoLine(strResponse)
				frmDrawMap.MsgQ.Add("You have no capital.")
				'UPGRADE_WARNING: Timer property frmDrawMap.Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
				If frmDrawMap.Msg_Timer.Interval = 0 Then
					frmDrawMap.Msg_Timer.Interval = 50
				End If
				
				'200304 rjk: Added to alert the user
			ElseIf InStr(strResponse, "You don't have the BTU's, bozo") > 0 Then 
				EchoLine(strResponse)
				frmDrawMap.MsgQ.Add("You don't have the BTU's, bozo.")
				If MsgBoxResult.Yes = MsgBox("You are out of BTU's" & vbLf & vbLf & "Do you want to logout and then login?", MsgBoxStyle.YesNo, "Out of BTU's") Then
					bRelogin = True
					WriteStringToSocket("bye" & vbCrLf)
				End If
				'UPGRADE_WARNING: Timer property frmDrawMap.Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
				If frmDrawMap.Msg_Timer.Interval = 0 Then
					frmDrawMap.Msg_Timer.Interval = 50
				End If
				
				'250304 rjk: Trap illegal commands
			ElseIf InStr(strResponse, "not a legal command") > 0 Then 
				EchoLine(strResponse)
				frmDrawMap.MsgQ.Add(strResponse)
				'UPGRADE_WARNING: Timer property frmDrawMap.Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
				If frmDrawMap.Msg_Timer.Interval = 0 Then
					frmDrawMap.Msg_Timer.Interval = 50
				End If
				'        NextCommand
				
				'Incoming mail announcement
			ElseIf (InStr(strResponse, "You have") > 0 And InStr(strResponse, "new telegram") > 0) And ExpectSpecificResponse <> READ_STATE_TELEGRAM_WRITE And ExpectSpecificResponse <> READ_STATE_FLASH_WRITE And ExpectSpecificResponse <> READ_STATE_WIRE And ExpectSpecificResponse <> READ_STATE_TELEGRAM_READ_CL And ExpectSpecificResponse <> READ_STATE_TELEGRAM_READ Then 
				
				If Not frmOptions.bSuppressTelegramNotification Then
					EchoLine(strResponse)
				End If
				'Make a sound
				If frmOptions.bSound Then '120203 rjk: Switched to frmOptions and boolean options.
					Beep()
				End If
				
				If bAutoRead Then
					'submit autoread
					SubmitTelegramRead(True, False)
					frmDrawMap.sb1.Items.Item("Mail").Text = " WAIT "
				Else
					'Set Mail indicator
					frmDrawMap.sb1.Items.Item("Mail").Text = " TELE "
				End If
				
				'get how many are pending
				'240104 rjk: Added spaces to match ??? and ...
				If frmDrawMap.sb1.Items.Item("Mail Count").Text <> " ??? " And frmDrawMap.sb1.Items.Item("Mail Count").Text <> " ... " Then
					If Len(frmDrawMap.sb1.Items.Item("Mail Count").Text) = 0 Then
						mailcount = 0
					Else
						mailcount = CShort(frmDrawMap.sb1.Items.Item("Mail Count").Text)
					End If
					n = ParseForNumber(strResponse)
					If n <= 0 Then
						frmDrawMap.sb1.Items.Item("Mail Count").Text = " ??? "
					Else
						mailcount = mailcount + n
						frmDrawMap.sb1.Items.Item("Mail Count").Text = " " & CStr(mailcount) & " "
					End If
				End If
				
				'check for new announcements
			ElseIf InStr(strResponse, "You have") > 0 And InStr(strResponse, "new announcement") > 0 And ExpectSpecificResponse <> READ_STATE_TELEGRAM_WRITE And ExpectSpecificResponse <> READ_STATE_FLASH_WRITE And ExpectSpecificResponse <> READ_STATE_WIRE And ExpectSpecificResponse <> READ_STATE_TELEGRAM_READ_CL And ExpectSpecificResponse <> READ_STATE_TELEGRAM_READ Then 
				EchoLine(strResponse)
				frmDrawMap.sb1.Items.Item("Anno").Text = " ANNO "
				
				'if pipes code, send message
			ElseIf nCode = EMP_C_PIPE Then 
				
				'piping generates an extra line, which can throw off line count routines
				'setting line to -1 conteracts this
				Line = -1
				MsgBox("WinACE does not pipe requests like " & strResponse)
				
			ElseIf nCode = EMP_C_EXECUTE Then 
				
				MsgBox("WinACE does not support the execute command. Instead, load " & "your batch file into the Script Editor Tool and run it from there." & "  Support for the execute command is targeted for the next release.")
				
				NextCommand(strResponse)
				'if redirect code - send to file
			ElseIf nCode = EMP_C_REDIR Then 
				'in case file open botchs
				On Error Resume Next
				Err.Clear()
				
				'redirection generates an extra line, which can throw off line count routines
				'setting line to -1 conteracts this
				Line = -1
				
				'test for redirect character
				n = InStr(strResponse, ">")
				If n > 0 Then
					OutputToFile = True
					bAppend = False
					n = n + 1
					'check for append
					If Mid(strResponse, n, 1) = ">" Then
						n = n + 1
						bAppend = True
					End If
					strFile = Trim(Mid(strResponse, n))
					
					'Append file suffix if necessary
					If InStr(strFile, ".") = 0 Then
						strFile = strFile & ".txt"
					End If
					If Not (VB.Left(strFile, 2) = "\\" Or Mid(strFile, 2, 1) = ":") Then
						strFile = My.Application.Info.DirectoryPath & "\" & strFile
					End If
					
					OutputFileNumber = FreeFile
					If bAppend Then
						FileOpen(OutputFileNumber, strFile, OpenMode.Append)
					Else
						FileOpen(OutputFileNumber, strFile, OpenMode.Output)
					End If
					If Len(Err.Description) > 0 Then
						OutputToFile = False
						EchoLine("Error accessing file " & strFile)
						EchoLine(Err.Description)
						frmDrawMap.lstCmdResult.Items.Add("Error accessing file " & strFile)
						frmDrawMap.lstCmdResult.Items.Add(Err.Description)
					Else
						EchoLine("Redirecting output to file " & strFile)
						frmDrawMap.lstCmdResult.Items.Add("Redirecting output to file " & strFile)
					End If
				End If
				On Error GoTo Empire_Error
				
				'if flash code - send to message box
			ElseIf nCode = EMP_C_FLASH Then 
				EchoLine(strResponse)
				'The server flashes when it aborts a command
				If InStr(strResponse, "Update aborting command") > 0 Then
					'Unload the query box if its up
					frmServerQuery.Close()
					
					MsgBox("Empire Server aborted command")
					
					'Reset all state of processing variables
					ExpectSpecificResponse = 0
					ServerQuery = False
					frmDrawMap.MsgQ.Add(strResponse)
					'100903 rjk: Added NextCommand = False to get out deadlock when doing offer treaty and an update abort comes.
					NextCommand()
				Else
					If frmOptions.bTTSFlashes And Len(strResponse) > 0 Then
						'                frmReport.ttsReport.Tag = "Normal"
						'                frmReport.ttsReport.Speak strResponse
					End If
					n = InStr(strResponse, "FLASH from")
					If n = 1 Then
						strTemp = Trim(Mid(strResponse, n + 10))
						n = InStr(strTemp, " @ ")
						If n > 0 Then
							If LastFlash <> Trim(VB.Left(strTemp, n)) Then
								LastFlash = Trim(VB.Left(strTemp, n))
							Else
								LastFlash = ""
							End If
							n = InStr(n, strTemp, ":")
							If n > 1 Then
								If InStr(n + 1, strTemp, ":") > 0 Then 'look for second colon
									n = InStr(n + 1, strTemp, ":")
									strTemp = Trim(Mid(strTemp, n + 1))
									If frmOptions.bFlashChat And Not (Me.frmFlash Is Nothing) Then
										Me.frmFlash.AddText(strResponse)
									End If
									FlashLog(strResponse)
								Else
									If InStr(n + 1, strTemp, "...") = n + 3 Then 'look for second colon
										n = InStr(n + 1, strTemp, "...")
										strTemp = Trim(Mid(strTemp, n + 3))
									Else
										strTemp = Trim(strTemp)
									End If
								End If
								frmDrawMap.MsgQ.Add("FLASH " & LastFlash & strTemp)
							Else
								frmDrawMap.MsgQ.Add(strResponse)
							End If
						Else
							frmDrawMap.MsgQ.Add(strResponse)
						End If
					Else
						n = InStr(strResponse, "BROADCAST from")
						If n = 1 Then
							strTemp = Trim(Mid(strResponse, n + 14))
							n = InStr(strTemp, " @ ")
							If n > 0 Then
								If LastFlash <> Trim(VB.Left(strTemp, n)) Then
									LastFlash = Trim(VB.Left(strTemp, n)) & " "
								Else
									LastFlash = ""
								End If
								n = InStr(n, strTemp, ":")
								If n > 0 Then
									If InStr(n + 1, strTemp, ":") > 0 Then 'look for second colon
										n = InStr(n + 1, strTemp, ":")
										strTemp = Trim(Mid(strTemp, n + 1))
										If frmOptions.bFlashChat And Not (Me.frmFlash Is Nothing) Then
											Me.frmFlash.AddText(strResponse)
										End If
										FlashLog(strResponse)
									Else
										If InStr(n + 1, strTemp, "...") = n + 3 Then 'look for second colon
											n = InStr(n + 1, strTemp, "...")
											strTemp = Trim(Mid(strTemp, n + 3))
										Else
											strTemp = Trim(strTemp)
										End If
									End If
									frmDrawMap.MsgQ.Add("BROADCAST " & LastFlash & strTemp)
								Else
									frmDrawMap.MsgQ.Add(strResponse)
								End If
							Else
								frmDrawMap.MsgQ.Add(strResponse)
							End If
						Else
							n = InStr(strResponse, "):")
							If n = 0 Then
								frmDrawMap.MsgQ.Add(strResponse)
								If InStr(strResponse, "LFLASH") = 1 Then
									If Me.bEmpireHub Then
										frmChat.AddText(strResponse)
									End If
								End If
							Else
								If InStr(strResponse, "<EOT>") = 0 Then
									If frmOptions.bFlashChat And Not (Me.frmFlash Is Nothing) Then
										Me.frmFlash.AddText(strResponse)
									End If
									FlashLog(strResponse)
								End If
								frmDrawMap.MsgQ.Add(Mid(strResponse, n + 3))
							End If
						End If
					End If
					'UPGRADE_WARNING: Timer property frmDrawMap.Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
					If frmDrawMap.Msg_Timer.Interval = 0 Then
						frmDrawMap.Msg_Timer.Interval = 50
					End If
				End If
				
				
				'if inform, then we need to submit a read
			ElseIf nCode = EMP_C_INFORM Then 
				If Len(strResponse) > 0 Then 'we seem to get bogus inform messages some times
					If frmOptions.bSound Then
						Beep() 'Make a sound
					End If
					If bAutoRead Then
						SubmitTelegramRead(True, False)
						frmDrawMap.sb1.Items.Item("Mail").Text = " ... "
						frmDrawMap.sb1.Items.Item("Mail Count").Text = " ... "
					Else
						If frmDrawMap.sb1.Items.Item("Mail Count").Text <> " ??? " And frmDrawMap.sb1.Items.Item("Mail Count").Text <> " ... " Then
							If Len(frmDrawMap.sb1.Items.Item("Mail Count").Text) = 0 Then
								mailcount = 0
							Else
								mailcount = CShort(frmDrawMap.sb1.Items.Item("Mail Count").Text)
							End If
							mailcount = mailcount + 1
							frmDrawMap.sb1.Items.Item("Mail Count").Text = " " & CStr(mailcount) & " "
						End If
						frmDrawMap.sb1.Items.Item("Mail").Text = " TELE "
					End If
				End If
			ElseIf nCode = EMP_C_HUB_TOKEN Then 
				If strResponse = "ToKeN" Then 'we seem to get bogus inform messages some times
					Select Case tokenStatus
						Case enumToken.tsSent
							If CmdQ.CmdsPending > 0 Then
								tokenStatus = enumToken.tsReceived
								ExecuteEmpireCmd(CmdQ.GetCommand)
							Else
								tokenStatus = enumToken.tsIdle
							End If
						Case Else
							Debug.Print("Received token response when not expecting one")
					End Select
				End If
				
			ElseIf ExpectSpecificResponse > 0 Then 
				
				
				Select Case ExpectSpecificResponse
					
					'This section is for when a sector dump has been requested
					Case READ_STATE_SECTOR_DUMP
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_HUB_COMMAND Then 
							NextCommand()
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							ExecuteEmpireCmd("*")
						Else
							Line = Line + 1
							
							'check for an announcement or tele message in the middle of a dump
							If InStr(strResponse, "You have") > 0 Then
								
								
							ElseIf Line > 3 Then 
								If InStr(strResponse, "sector") > 0 Then
									'We are done - this is sector count
									Line = 0
									EchoLine(strResponse)
									ExpectSpecificResponse = 0
								Else
									If SubmittedFromCommandLine Then
										EchoLine(strResponse)
									End If
									If Len(ProcessSectorDump(strResponse, sx, sy)) > 0 Then
										'update the sector box if necessary
										If sx = frmDrawMap.SelX And sy = frmDrawMap.SelY Then
											frmDrawMap.FillSectorBox(sx, sy)
										End If
									End If
								End If
							ElseIf Line = 2 And UpdateTimeStamp Then 
								If InStr(strResponse, "DUMP") > 0 Then
									strResponse = Trim(strResponse)
									n = InStr(strResponse, " ")
									n = InStr(n + 1, strResponse, " ")
									tsSectors = CInt(Trim(Mid(strResponse, n + 1)))
								End If
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							Else
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							End If
						End If
						
						'This handles parsing the nation command
					Case READ_STATE_STRENGTH
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress Then
								If Not frmReport.Visible Then ShowReportBox()
							End If
							NextCommand(strResponse)
						Else
							Line = Line + 1
							'Add to list boxes
							If Not DatabaseUpdateinProgress Then
								frmReport.AddReportLine(strResponse)
								EchoLine(strResponse)
							End If
							'102103 rjk: Need to process an empty set, no sectors.
							'111603 rjk: Added not a legal command for sanutary startup.
							'If InStr(strResponse, "sector") > 0 And Line > 3 Then
							If (InStr(strResponse, "sector") > 0 And Line > 3) Or InStr(strResponse, "No sector(s)") > 0 Or InStr(strResponse, "not a legal command") > 0 Then
								'We are done - this is sector count
								Line = 0
								If DatabaseUpdateinProgress Then
									EchoLine(strResponse)
								End If
								ExpectSpecificResponse = 0
								If Not DatabaseUpdateinProgress Then
									If Not frmReport.Visible Then ShowReportBox()
								End If
							ElseIf Line > 3 Then 
								ParseStrength(strResponse)
							End If
						End If
						
						'This section is for when a lost dump has been requested
					Case READ_STATE_LOST_DUMP
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_HUB_COMMAND Then 
							NextCommand()
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							ExecuteEmpireCmd("*")
						Else
							Line = Line + 1
							
							If Line = 1 Then
								rsSectors.Index = "PrimaryKey"
								rsShip.Index = "PrimaryKey"
								rsPlane.Index = "PrimaryKey"
								rsLand.Index = "PrimaryKey"
								rsNuke.Index = "PrimaryKey"
							ElseIf Line > 3 Then 
								If InStr(strResponse, "lost") > 0 Then
									'We are done - this is sector count
									Line = 0
									EchoLine(strResponse)
									ExpectSpecificResponse = 0
								Else
									If SubmittedFromCommandLine Then
										EchoLine(strResponse)
									End If
									ProcessLostDump(strResponse)
								End If
							ElseIf Line = 2 And UpdateTimeStamp Then 
								If InStr(strResponse, "DUMP") > 0 Then
									strResponse = Trim(strResponse)
									n = InStr(strResponse, " ")
									n = InStr(n + 1, strResponse, " ")
									n = InStr(n + 1, strResponse, " ")
									If n <> 0 Then '101503 rjk: prevents run-time errors if the dumps get out of step
										tsLost = CInt(Trim(Mid(strResponse, n + 1)))
									End If
								End If
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							Else
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							End If
						End If
						
						'This section is for when a Ship dump has been requested
					Case READ_STATE_SHIP_DUMP
						ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_SHIP_DUMP, nCode, Line, rsShip, strResponse, "ship", HaveShips, tsShip)
						
					Case READ_STATE_PLANE_DUMP
						ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_PLANE_DUMP, nCode, Line, rsPlane, strResponse, "plane", HavePlanes, tsPlane)
						
						
					Case READ_STATE_LAND_DUMP
						ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_LAND_DUMP, nCode, Line, rsLand, strResponse, "unit", HaveLands, tsLand)
						
						
						'nukes are just slightly different than other units, so we'll leave them here
						'instead of factoring out like the above.
					Case READ_STATE_NUKE_DUMP
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							ExecuteEmpireCmd("*")
						Else
							Line = Line + 1
							'101403 rjk: Modified ProcessNukeDump to work with partial dumps and ProcessLostDump
							'            will cleanup extra records.
							'If Line = 1 Then
							'    If Not (rsNuke.EOF And rsNuke.BOF) Then
							'        'Clear nuke file
							'        rsNuke.MoveFirst
							'        While Not (rsNuke.EOF And rsNuke.BOF)
							'            rsNuke.Delete
							'            rsNuke.MoveFirst
							'        Wend
							'    End If
							'ElseIf Line > 3 Then
							If Line > 3 Then
								n = InStr(strResponse, "nuke")
								If n > 0 Then
									'We are done - this is nuke count
									If rsNuke.BOF And rsNuke.EOF Then
										HaveNukes = False
									Else
										HaveNukes = True
									End If
									
									Line = 0
									EchoLine(strResponse)
									ExpectSpecificResponse = 0
								Else
									If SubmittedFromCommandLine Then
										EchoLine(strResponse)
									End If
									ProcessNukeDump(strResponse, rsNuke)
								End If
							ElseIf Line = 2 And UpdateTimeStamp Then 
								If InStr(strResponse, "DUMP") > 0 Then
									strResponse = Trim(strResponse)
									n = InStr(strResponse, " ")
									n = InStr(n + 1, strResponse, " ")
									tsNuke = CInt(Trim(Mid(strResponse, n + 1)))
								End If
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							Else
								If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
									EchoLine(strResponse)
								End If
							End If
						End If
						
						'This section is for when a full Bmap has been requested
					Case READ_STATE_FULL_BMAP
						If nCode = EMP_C_PROMPT Then
							
							'This code checks for an empty bmap.  If found, it runs a
							'regular map.  This should only be necessary before breaking
							'sanctuary
							If rsBmap.BOF And rsBmap.EOF And BmapFirstTimeFlag = 0 Then
								Me.SubmitEmpireCommand("db1", False)
								Me.SubmitEmpireCommand("map #1", False)
								Me.SubmitEmpireCommand("db2", False)
								BmapFirstTimeFlag = 999
							End If
							
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							ExecuteEmpireCmd("*")
						Else
							Line = Line + 1
							
							
							If Line = 1 Then
								bmaphead1 = strResponse
								linelen = Len(RTrim(bmaphead1))
							ElseIf Line = 2 Then 
								bmaphead2 = strResponse
								'Calculate the starting x
								StartX = Val(Mid(bmaphead1, 6, 1)) * 10 + Val(Mid(bmaphead2, 6, 1))
								temp = Val(Mid(bmaphead1, 7, 1)) * 10 + Val(Mid(bmaphead2, 7, 1))
								If temp < StartX Then
									StartX = -1 * StartX
								End If
							ElseIf Line = 3 Then 
								If VB.Left(strResponse, 4) = "    " Then
									'Handle three line bmap - recompute start x
									bmaphead3 = strResponse
									StartX = Val(Mid(bmaphead1, 6, 1)) * 100 + Val(Mid(bmaphead2, 6, 1)) * 10 + Val(Mid(bmaphead3, 6, 1))
									temp = Val(Mid(bmaphead1, 7, 1)) * 100 + Val(Mid(bmaphead2, 7, 1)) * 10 + Val(Mid(bmaphead3, 7, 1))
									If temp < StartX Then
										StartX = -1 * StartX
									End If
								Else
									'handle special case from Big Game
									'World Size 200 generates 00 as starting -x, but no third line
									If MapSizeX = 200 And StartX = 0 Then
										StartX = -100
									End If
									LoadBmap(strResponse, StartX, linelen)
								End If
							Else
								LoadBmap(strResponse, StartX, linelen)
							End If
						End If
						
					Case READ_STATE_XDUMP '270804 rjk: Added to Parse xdump
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_HUB_COMMAND Then 
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							If DisplayCmd Then
								EchoLine(strResponse)
							End If
							If Line = -1 Then
								Line = 0
							Else
								ParseXdump(strResponse, UpdateTimeStamp)
							End If
						End If
						
						'This is for reports that go to the report dialog box
					Case READ_STATE_REPORT, READ_STATE_WIRE
						If nCode = EMP_C_PROMPT Then
							'drk@unxsoft.com 2/26/03
							If Not frmReport.Visible Then ShowReportBox()
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If Len(strReply1) > 0 Then
								EchoLine(strResponse & strReply1)
								ServerQuery = True
								SendStringtoServer(strReply1)
								strReply1 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							frmReport.AddReportLine(strResponse)
							EchoLine(strResponse)
						End If
						
						'This is for the inital report request to build nation list
						'100703 rjk: READ_STATE_REQ_REPORT not required a nation list is done in relations now.
						'112503 rjk: Added Report back in.
						'Case READ_STATE_REQ_REPORT, READ_STATE_REQ_RELATIONS
					Case READ_STATE_REQ_RELATIONS
						'121503 rjk: If the whole Relation response is not processed in ProcessServerResponse
						'            NationFrom was forgotten.  Problem with EmpireHub.
						
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress Then
								If Not frmReport.Visible Then ShowReportBox() 'drk 2/26/03
							End If
							NextCommand(strResponse)
							Line = 0
							frmChat.ControlFlashForm()
						Else
							Line = Line + 1
							If Not DatabaseUpdateinProgress Then '121203 rjk: Added to display relations output if done on command line or from the menu.
								frmReport.AddReportLine(strResponse)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							'100703 rjk: move logic for creating rsNation and Nation list to Relations Report
							'            first report must be relation of home country to fill grid
							'121703 rjk: Added a check for Diplomatic Relations title to prevent crash with invalid syntax
							If Line = 1 And InStr(1, strResponse, "Diplomatic Relations Report") > 3 Then
								strResponse = Trim(strResponse)
								'find out who this relations report is FROM.  note that you only get
								'this type of header if you've specified a nation NUMBER on the cmd line.
								'if you specify a nation name, it doesn't display it.
								'100703 rjk: Space not a character to look, names can have spaces
								'NationFrom = Nations.NationNumber(Mid$(strResponse, 2, InStr(1, strResponse, " ") - 2))
								NationFromName = Mid(strResponse, 2, InStr(1, strResponse, "Diplomatic Relations Report") - 3)
								NationFrom = Nations.NationNumber(NationFromName)
								If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
									If CountryName = NationFromName Or CountryNumber = NationFrom Then
										'clear the nation database
										DeleteAllRecords(rsNations)
										'add ourselves
										Nations.AddNation(CountryNumber, NationFromName)
										rsNations.AddNew()
										rsNations.Fields("Name").Value = CountryName
										rsNations.Fields("Number").Value = CountryNumber
										rsNations.Update()
										CountryName = NationFromName 'If our name changes adapt
										NationFrom = CountryNumber
										On Error GoTo Empire_Error
									End If
								End If
							ElseIf Mid(strResponse, 4, 1) = ")" Then 
								NationTowards = CShort(Trim(VB.Left(strResponse, 3)))
								
								'Ensure the NationTowards is in the country list
								'This is important with the hidden option as this list
								'is another source of countries name
								'NationName = Trim$(Mid$(strResponse, 6, 16))
								'100703 rjk: Should only need to add we are on our own relations (done first)
								If CountryNumber = NationFrom And VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
									NationName = Trim(Mid(strResponse, 6, 19))
									Nations.AddNation(NationTowards, NationName)
									rsNations.AddNew()
									rsNations.Fields("Name").Value = NationName
									rsNations.Fields("Number").Value = NationTowards
									rsNations.Update()
								End If
								
								NationRel = Nations.RelationIndex(Mid(strResponse, 27, 11))
								
								'update the local table
								Nations.Relation(NationFrom, NationTowards) = NationRel
							End If
						End If
						
						'this is for when we break sanctuary
					Case READ_STATE_BREAK
						If nCode = EMP_C_PROMPT Then
							'Update the database
							DeleteAllRecords(rsShipOrders)
							DeleteAllRecords(rsMissions)
							DeleteAllRecords(rsSea)
							DeleteAllRecords(rsShip)
							DeleteAllRecords(rsLand)
							DeleteAllRecords(rsNuke)
							DeleteAllRecords(rsPlane)
							frmTelegram.ClearTelegrams()
							ClearEnemyInfo(True, True, True, True, True, True, True, True, 0) '112903 rjk: Added individual control to what is deleted
							SubmitEmpireCommand("db1", False)
							If Not (VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS) Then
								RequestMetaTables()
								RequestConfigurationTables()
							End If
							If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS And Not frmOptions.bSPAtlantis Then
								GetNation()
								GetCountryList()
							Else
								GetCountryList()
								GetNation()
							End If
							GetSectors(True)
							'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
							SubmitEmpireCommand("strength *", False)
							'Need to do a map, since a bmap has not been updated yet
							SubmitEmpireCommand("map #1", False)
							SubmitEmpireCommand("db2", False)
							'toggle inform/flash/techlist on
							SubmitEmpireCommand("toggle techlist on")
							SubmitEmpireCommand("toggle flash on")
							SubmitEmpireCommand("toggle inform on")
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							EchoLine(strResponse)
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							NewLineCount = NewLineCount + 1
						End If
						
						'This is for multiple explores in a circle
					Case READ_STATE_EXPLORE_CIRCLE
						
						If nCode = EMP_C_PROMPT Then
							'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
							frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							'check for abandon query, or risk loop
							If InStr(strResponse, "abandon") = 0 Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse & "h")
								EchoLine(strResponse & "h")
								ServerQuery = True
								SendStringtoServer("h")
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							EchoLine(strResponse)
						End If
						
						
						'This is unit moves (nav, march)
					Case READ_STATE_UNIT_MOVE
						If nCode = EMP_C_PROMPT Then
							If AutoMoveinProgress Then
								ShowReportBox()
							Else
								frmNavCtrl.Close() 'Close the box
								If bUnitMoveShowReport Then
									ShowReportBox()
									bUnitMoveShowReport = False
								End If
							End If
							AutoMoveinProgress = False
							AutoMoveLook = False
							AutoMoveView = False
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If InStr(strResponse, "abandon") > 0 Then
								RespondtoServerQuery(strResponse)
							ElseIf AutoMoveinProgress Then 
								If AutoMoveLook Or AutoMoveView Then
									If AutoMoveLook Then
										frmReport.AddReportLine(strResponse & "l")
										EchoLine(strResponse & "l")
										ServerQuery = True
										SendStringtoServer("l")
										AutoMoveLook = False
										Looking = True
										'                            End If
									ElseIf AutoMoveView Then 
										frmReport.AddReportLine(strResponse & "v")
										EchoLine(strResponse & "v")
										ServerQuery = True
										SendStringtoServer("v")
										AutoMoveView = False
										Looking = False
									End If
								Else
									Looking = False
									frmReport.AddReportLine(strResponse & "h")
									EchoLine(strResponse & "h")
									ServerQuery = True
									SendStringtoServer("h")
								End If
							Else
								
								frmNavCtrl.AddReportLine(strResponse)
								ServerQuery = True
								frmNavCtrl.Looking = False
								frmNavCtrl.Sonar = False
								'Position list box
								'UPGRADE_ISSUE: ListBox property lstReport.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
								frmNavCtrl.lstReport.TopIndex = frmNavCtrl.lstReport.NewIndex
								
								frmNavCtrl.EnableButtons(True)
								
								If Not frmNavCtrl.Visible Then
									frmNavCtrl.Show()
								End If
								
								If frmNavCtrl.AutoLooking_when_notAutoStopping Or frmNavCtrl.AutoViewing_when_notAutoStopping Then frmNavCtrl.callback()
							End If
							
						Else
							EchoLine(strResponse)
							'check for a sunk ship (interdiction, mines, etc)
							If InStr(strResponse, "sunk!") > 0 Then
								BulletinShipSunk(strResponse)
								If Not SubmitedfromCmdWindow Then
									bUnitMoveShowReport = True
								End If
							ElseIf (InStr(strResponse, "spotted") > 0 Or InStr(strResponse, "fires at you") > 0 Or InStr(strResponse, "interdiction") > 0 Or InStr(strResponse, "has no mil on it to guide ") > 0 Or InStr(strResponse, "is out of mobility & stays in") > 0 Or InStr(strResponse, "is crewless & stays in") > 0) And Not SubmitedfromCmdWindow Then 
								bUnitMoveShowReport = True
								'check for a view output showing fert or oil
							ElseIf InStr(strResponse, "[oil") > 0 Or InStr(strResponse, "[fert") > 0 Then 
								ParseShipView(strResponse)
							End If
							
							If AutoMoveinProgress Then
								'Put Line in Listbox
								frmReport.AddReportLine(strResponse)
								If Looking And Len(strResponse) > 10 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ParseLook(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									VARsec = ParseLook(strResponse)
									If IsArray(VARsec) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If VARsec(0) = "h" Then
											AddEnemySectInfo(VARsec)
											'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ElseIf VARsec(0) = "l" Or VARsec(0) = "s" Or VARsec(0) = "p" Then 
											AddEnemyUnitInfo(VARsec)
										End If
									End If
								End If
							ElseIf Not DatabaseUpdateinProgress Then 
								'Put the line in the Nav Control form
								frmNavCtrl.AddReportLine(strResponse)
								If Not frmNavCtrl.Visible Then
									frmNavCtrl.Show()
								End If
								If bUnitMoveShowReport Then
									frmReport.AddReportLine(strResponse)
								End If
								'if this was a look command - parse it
								'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								VARsec = 0
								If frmNavCtrl.Looking And Len(strResponse) > 10 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ParseLook(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									VARsec = ParseLook(strResponse)
								End If
								'060304 rjk: Added Sonar parsing to Navigation form.
								If frmNavCtrl.Sonar And InStr(strResponse, "Sonar detects") > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ParseSonar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									VARsec = ParseSonar(Trim(Mid(strResponse, 14)))
								End If
								If IsArray(VARsec) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If VARsec(0) = "h" Then
										AddEnemySectInfo(VARsec)
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ElseIf VARsec(0) = "l" Or VARsec(0) = "s" Or VARsec(0) = "p" Then 
										AddEnemyUnitInfo(VARsec)
									End If
								End If
								
							End If
						End If
						
						'This is for next update
					Case READ_STATE_NEXT_UPDATE
						
						
						If nCode = EMP_C_PROMPT Then
							
							'see below. drk@unxsoft.com 6/30/01
							If TimeUpdate = "demand" Then
								frmDrawMap.UpdateTimer.Enabled = False
								frmDrawMap.sb1.Items.Item("Timer").Text = "Demand"
							Else
								'102103 rjk: Added to check in case the update time is missing
								If Len(TimeUpdate) > 0 Then
									timNextUpdate = EmpireTimeString("    " & TimeUpdate & " " & CStr(Year(Now)))
									'UPGRADE_WARNING: Couldn't resolve default property of object SecondsDifference(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									frmDrawMap.SecondsToUpdate = SecondsDifference(TimeCurrent, TimeUpdate)
									sec = frmDrawMap.SecondsToUpdate Mod 60
									totmin = (frmDrawMap.SecondsToUpdate - sec) \ 60
									If totmin <= 59 Then
										Min = totmin
										hour_Renamed = 0
									Else
										Min = totmin Mod 60
										hour_Renamed = totmin \ 60
										If hour_Renamed > 99 Then
											hour_Renamed = 99
										End If
									End If
									frmDrawMap.sb1.Items.Item("Timer").Text = " " & VB6.Format(hour_Renamed, "00") & ":" & VB6.Format(Min, "00") & ":" & VB6.Format(sec, "00") & " "
									frmDrawMap.UpdateTimer.Interval = 1000
									frmDrawMap.UpdateTimer.Enabled = True
								Else
									'102003 rjk: timNextUpdate is used and AirCombat History and may not work correctly
									timNextUpdate = EmpireTimeString("    " & TimeCurrent & " " & CStr(Year(Now)))
									frmDrawMap.sb1.Items.Item("Timer").Text = "N/A"
									frmDrawMap.UpdateTimer.Enabled = False
								End If
								If rsVersion.BOF And rsVersion.EOF Then
									rsVersion.AddNew()
								Else
									rsVersion.MoveFirst()
									rsVersion.Edit()
								End If
								'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
								rsVersion.Fields("Time Zone Adj").Value = DateDiff(Microsoft.VisualBasic.DateInterval.Minute, Now, EmpireTimeString("    " & TimeCurrent & " " & CStr(Year(Now))))
								rsVersion.Update()
								rsVersion.MoveFirst()
							End If
							
							NextCommand(strResponse)
						Else
							'Put repsonse into list box
							EchoLine(strResponse)
							If DisplayCmd Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							
							'drk@unxsoft.com 6/30/01 : bug fix: we got really confused on demand updates only
							'(which, I can imagine is not a very popular option for real games, but I
							'was trying to use it on a local server to test my development)
							If strResponse = "There are no regularly scheduled updates." Then
								'demand updates only.
								TimeUpdate = "demand"
							Else
								n = InStr(strResponse, "next update is at")
								frmDrawMap.TimerAtUpdate = VB.Timer()
								If n > 0 Then
									TimeUpdate = Trim(Mid(strResponse, 27, Len(strResponse) - 27))
								Else
									n = InStr(strResponse, "current time is")
									If n > 0 Then
										TimeCurrent = Trim(Mid(strResponse, n + 15))
										TimeCurrent = Mid(TimeCurrent, 5, Len(TimeCurrent) - 5)
									End If
								End If
							End If
						End If
						
						'This is for multiple fires in sequence
					Case READ_STATE_FIRE_SEQ
						If nCode = EMP_C_PROMPT Then
							If BatchFileinProgress Then
								'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
								frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
								NextCommand(strResponse)
							Else
								NextCommand(strResponse)
							End If
							
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							If InStr(strResponse, "sunk!") > 0 Then
								BulletinShipSunk(strResponse)
							End If
							'Put repsonse into list box
							EchoLine(strResponse)
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							NewLineCount = NewLineCount + 1
							
							'Shell hit sector -3,-1 for 40 damage.
							If InStr(strResponse, "Shell hit sector") > 0 Or InStr(strResponse, "Shells hit sector") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "sector", "for"))
								shelldamage = ParseShellDamage(strResponse)
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
								NewLineCount = NewLineCount + 1
							End If
							
							'if return fire, set batch fire to abort
							If InStr(strResponse, "Defenders fire back!") > 0 Then
								If BatchFireAbort > BF_NEVERABORT Then
									BatchFireAbort = BF_ABORTREMAINING
								End If
							End If
							
							
							
							'lc   light cruiser (#32) takes 55
							If InStr(strResponse, "(#") > 0 And InStr(strResponse, " takes ") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "(#", ")"))
								shelldamage = CShort(Trim(StringBetween(strResponse, "takes")))
								strResponse = ApplyShipDamage(shellsector, shelldamage, True)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
								NewLineCount = NewLineCount + 1
							End If
							
							'Torpedo hit frg  frigate (#1) for 75 damage.
							'Shell hit frg  frigate (#1) in -8,-8 for 38 damage.
							'Shells hit frg  frigate (#1) in -8,-8 for 107 damage.
							
						End If
						
						'This is for launch
					Case READ_STATE_LAUNCH
						If nCode = EMP_C_PROMPT Then
							ShowReportBox()
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If Len(strReply1) > 0 Then
								frmReport.AddReportLine(strResponse & strReply1)
								EchoLine(strResponse & strReply1)
								ServerQuery = True
								SendStringtoServer(strReply1)
								strReply1 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							If InStr(strResponse, "sunk!") > 0 Then
								BulletinShipSunk(strResponse)
							End If
							
							frmReport.AddReportLine(strResponse)
							EchoLine(strResponse)
							
							'did 42 damage in 12,-4
							If InStr(strResponse, " damage in ") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "in "))
								If InStr(strResponse, "%") > 0 Then
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "%")))
								Else
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "damage")))
								End If
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmReport.AddReportLine(strResponse)
							End If
							
						End If
						
						'This is for bombing runs
					Case READ_STATE_BOMB
						If nCode = EMP_C_PROMPT Then
							ShowReportBox()
							strReply1 = bombreply1
							strReply2 = bombreply2
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If InStr(strResponse, "effici") > 0 And Len(strReply1) > 0 Or InStr(strResponse, "commod") > 0 And Len(strReply1) Then
								frmReport.AddReportLine(strResponse & strReply1)
								ServerQuery = True
								EchoLine(strResponse & strReply1)
								SendStringtoServer(strReply1)
								
								'cache reply for multiple bombing
								strReply1 = vbNullString
							ElseIf frmOptions.bSPAtlantis And Len(strReply1) > 0 And (InStr(strResponse, "eff") > 0 Or InStr(strResponse, "comms") > 0 Or InStr(strResponse, "infra") > 0) Then 
								frmReport.AddReportLine(strResponse & strReply1)
								ServerQuery = True
								EchoLine(strResponse & strReply1)
								SendStringtoServer(strReply1)
								
								'cache reply for multiple bombing
								strReply1 = vbNullString
							ElseIf (InStr(strResponse, "bomb?") > 0 Or InStr(strResponse, "Target?") > 0) And Len(strReply2) > 0 Then 
								frmReport.AddReportLine(strResponse & strReply2)
								EchoLine(strResponse & strReply2)
								ServerQuery = True
								SendStringtoServer(strReply2)
								'cache reply for multiple bombing
								strReply2 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							If Not SubmittedFromCommandLine Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							frmReport.AddReportLine(strResponse)
							EchoLine(strResponse)
							If InStr(strResponse, "flying over ") > 0 Then
								BulletinFlewOver(strResponse)
							End If
							If InStr(strResponse, "sunk!") > 0 Then
								BulletinShipSunk(strResponse)
							End If
							'check for carrier landing - "landing on carrier 0"
							If InStr(strResponse, "landing on carrier") > 0 And Not CarrierDumpRequested Then
								CarrierDumpRequested = True
								GetShips()
							End If
							
							'did 59% damage to efficiency in 1,-3
							If InStr(strResponse, "damage to efficiency") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "in "))
								shelldamage = CShort(Trim(StringBetween(strResponse, "did", "%")))
								strResponse = ApplySectorDamage(shellsector, shelldamage, True)
								EchoLine(strResponse)
								frmReport.AddReportLine(strResponse)
							End If
							
							'1,-3 takes 14% collateral damage
							If InStr(strResponse, "collateral damage") > 0 Then
								shellsector = Trim(VB.Left(strResponse, InStr(strResponse, "takes ") - 1))
								shelldamage = CShort(Trim(StringBetween(strResponse, "takes", "%")))
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmReport.AddReportLine(strResponse)
							End If
							
							'did 31 damage in 15,-3
							If InStr(strResponse, " damage in ") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "in "))
								If InStr(strResponse, "%") > 0 Then
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "%")))
								Else
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "damage")))
								End If
								shelldamage = CShort(shelldamage / (1# - (CSng(shelldamage) / 100)))
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmReport.AddReportLine(strResponse)
							End If
							
						End If
						
						'This is for bombing runs from command line
					Case READ_STATE_BOMBCL
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							EchoLine(strResponse)
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							If InStr(strResponse, "flying over ") > 0 Then
								BulletinFlewOver(strResponse)
							End If
							
							If InStr(strResponse, "sunk!") > 0 Then
								BulletinShipSunk(strResponse)
							End If
							
							'did 59% damage to efficiency in 1,-3
							If InStr(strResponse, "damage to efficiency") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "in "))
								shelldamage = CShort(Trim(StringBetween(strResponse, "did", "%")))
								strResponse = ApplySectorDamage(shellsector, shelldamage, True)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							
							'1,-3 takes 14% collateral damage
							If InStr(strResponse, "collateral damage") > 0 Then
								shellsector = Trim(VB.Left(strResponse, InStr(strResponse, "takes ") - 1))
								shelldamage = CShort(Trim(StringBetween(strResponse, "takes", "%")))
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							
							'did 31 damage in 15,-3
							'did 31%  damage in 15,-3 120404 rjk: Added for processing nuke output
							If InStr(strResponse, " damage in ") > 0 Then
								shellsector = Trim(StringBetween(strResponse, "in "))
								If InStr(StringBetween(strResponse, "did", "damage"), "%") > 0 Then
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "% damage")))
								Else
									shelldamage = CShort(Trim(StringBetween(strResponse, "did", "damage")))
								End If
								shelldamage = CShort(shelldamage / (1# - (CSng(shelldamage) / 100)))
								strResponse = ApplySectorDamage(shellsector, shelldamage)
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							
						End If
						
						'This is for read
					Case READ_STATE_TELEGRAM_READ
						If nCode = EMP_C_PROMPT Then
							'Set Mail picture
							frmDrawMap.sb1.Items.Item("Mail").Text = " "
							frmDrawMap.sb1.Items.Item("Mail Count").Text = vbNullString 'drk@unxsoft.com 5/8/03.  this was set to " " before, which caused crashes when it was converted to an Int elsewhere in this routine (the next time through, I think).
							If RecievedNoticeOfMapChange Then
								SubmitEmpireCommand("db2", False) 'Update map
								RecievedNoticeOfMapChange = False
							End If
							LastFlash = vbNullString 'turn off so flash header rewrites
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If Len(strReply1) > 0 Then
								EchoLine(strResponse & strReply1)
								ServerQuery = True
								SendStringtoServer(strReply1)
								strReply1 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							If InStr(strResponse, "No telegrams for you") = 0 Then
								IncomingTelegramLine(strResponse, VB.Left(strResponse, 1) = ">" And InStr(strResponse, "  dated ") > 0)
								If Not UnitDamage(strResponse) Or Not frmOptions.bSuppressUnitDamageReports Then
									frmDrawMap.MsgQ.Add(TelegramMessage)
								End If
								'UPGRADE_WARNING: Timer property frmDrawMap.Msg_Timer.Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
								If frmDrawMap.Msg_Timer.Interval = 0 Then
									frmDrawMap.Msg_Timer.Interval = 50
								End If
								
							End If
							If Not UnitDamage(strResponse) Or Not frmOptions.bSuppressUnitDamageReports Then
								EchoLine(strResponse)
							End If
						End If
						
						'This is for telegram/announcement
					Case READ_STATE_TELEGRAM_WRITE, READ_STATE_FLASH_WRITE
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If Len(strReply1) > 0 Then
								n = InStr(strReply1, vbLf)
								If n = 0 Then n = Len(strReply1) + 1
								ServerQuery = True
								EchoLine(strResponse & VB.Left(strReply1, n - 1))
								SendStringtoServer(VB.Left(strReply1, n - 1))
								strReply1 = Mid(strReply1, n + 1)
							ElseIf Len(strReply2) > 0 Then 
								EchoLine(strResponse & strReply2)
								ServerQuery = True
								SendStringtoServer(strReply2)
								strReply2 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						End If
						
					Case READ_STATE_SCUTTLE '040404 rjk: Added for automatic answer for scuttle and scrap
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If Len(strReply1) > 0 Then
								ServerQuery = True
								SendStringtoServer(strReply1)
								strReply1 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							EchoLine(strResponse)
							If DisplayCmd Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
						End If
						
						'This is look reports from ships and lands
					Case READ_STATE_LOOK, READ_STATE_SPY, READ_STATE_RECON, READ_STATE_COASTWATCH, READ_STATE_SONAR
						If nCode = EMP_C_PROMPT Then
							AntiSubPatrolinProgress = False
							ShowReportBox()
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							frmReport.AddReportLine(strResponse)
							If DisplayCmd Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							Select Case ExpectSpecificResponse
								Case READ_STATE_SPY
									'UPGRADE_WARNING: Couldn't resolve default property of object ParseSpyReport(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									VARsec = ParseSpyReport(strResponse)
								Case READ_STATE_RECON
									'Check for Anti Sub Patrol
									If AntiSubPatrolinProgress Then
										'sb   submarine (#12) 6,4
										'WinACE2 sb   submarine (#13) @ 6,4
										'sub #1 8,0
										'101503 rjk: Added processing for land sectors as well
										'If InStr(strResponse, " (#") > 0 And InStr(strResponse, "carrier") = 0 Then  'sub found
										If InStr(strResponse, " (#") > 0 Then
											If InStr(strResponse, "efficient") > 0 Then 'land sector
												'UPGRADE_WARNING: Couldn't resolve default property of object ParseReconReport(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												VARsec = ParseReconReport(strResponse)
											ElseIf InStr(strResponse, "carrier") = 0 Then  'sub
												If InStr(strResponse, " downgraded ") > 0 Then '123103 rjk: Parse an relations change during a np recon
													'Country WinACEdv2 (#4) has downgraded their relations with you to "Hostile"!
													'Another cold war...
													'Diplomatic relations with WinACEdv2 downgraded to "Hostile".
													'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													VARsec = 0
												ElseIf InStr(strResponse, " @ ") = 0 Then  'special missing @
													n = InStr(strResponse, ") ")
													'UPGRADE_WARNING: Couldn't resolve default property of object ParseSonar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													VARsec = ParseSonar(Trim(VB.Left(strResponse, n + 1) & "@ " & Mid(strResponse, n + 2)))
												Else
													'UPGRADE_WARNING: Couldn't resolve default property of object ParseSonar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													VARsec = ParseSonar(Trim(strResponse))
												End If
											End If
										ElseIf VB.Left(strResponse, 5) = "sub #" Then 
											n = InStr(6, strResponse, " ")
											strTemp = "sb   submarine (#" & Mid(strResponse, 6, n - 6) & ") @ " & Mid(strResponse, n + 1)
											'UPGRADE_WARNING: Couldn't resolve default property of object ParseSonar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											VARsec = ParseSonar(Trim(strTemp))
										End If
									ElseIf InStr(strResponse, "Anti-Sub Patrol") > 0 Then 
										'Anti-Sub Patrol report
										AntiSubPatrolinProgress = True
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object ParseReconReport(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										VARsec = ParseReconReport(strResponse)
									End If
								Case READ_STATE_LOOK
									If InStr(strResponse, "Usage") = 0 Then '120503 rjk: Skip an invalid llook command
										'UPGRADE_WARNING: Couldn't resolve default property of object ParseLook(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										VARsec = ParseLook(strResponse)
									End If
								Case READ_STATE_COASTWATCH
									'UPGRADE_WARNING: Couldn't resolve default property of object ParseCoastWatch(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									VARsec = ParseCoastWatch(strResponse)
								Case READ_STATE_SONAR
									'Sonar detects Furrym sb   submarine (#956) @ -31,29
									If VB.Left(strResponse, 13) = "Sonar detects" Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ParseSonar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										VARsec = ParseSonar(Trim(Mid(strResponse, 14)))
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										VARsec = 0
									End If
									
							End Select
							If IsArray(VARsec) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If VARsec(0) = "h" Then
									AddEnemySectInfo(VARsec)
									'UPGRADE_WARNING: Couldn't resolve default property of object VARsec(0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ElseIf VARsec(0) = "l" Or VARsec(0) = "s" Or VARsec(0) = "p" Then 
									AddEnemyUnitInfo(VARsec)
								End If
							End If
							
							EchoLine(strResponse)
						End If
						
						'this handles starvation, paradrop reports
					Case READ_STATE_STARVATION, READ_STATE_PARA
						If nCode = EMP_C_PROMPT Then
							If Not frmReport.Visible Then ShowReportBox()
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							frmReport.AddReportLine(strResponse)
							EchoLine(strResponse)
							If DisplayCmd Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							' add a marker
							Line = Line + 1
							If ExpectSpecificResponse = READ_STATE_PARA Then
								If InStr(strResponse, "flying over ") > 0 Then
									BulletinFlewOver(strResponse)
								End If
							End If
							If ExpectSpecificResponse = READ_STATE_STARVATION Then
								If Line = 1 Then
									EventMarkers.Clear()
								Else
									If InStr(strResponse, ",") Then
										EventStarvationNotice(strResponse)
									End If
								End If
							End If
						End If
						
						'this handles starvation reports
					Case READ_STATE_SHOW
						
						If nCode = EMP_C_PROMPT Then
							'dummy call to parseshow to make sure first line is updated
							If Len(ReportType) > 0 And Len(UnitType) > 0 Then
								ParseShow(vbNullString, 1, UnitType, ReportType, 0)
								
								If UnitType = "c" And (ReportType = "s" Or ReportType = "c") Then
									LoadSectorDesc()
								End If
								If frmToolShow.Visible Then
									frmToolShow.FillGrid()
								End If
							End If
							NextCommand(strResponse)
							Line = 0
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							EchoLine(strResponse)
							If DisplayCmd Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							' add a marker
							Line = Line + 1
							Select Case Line
								Case 1
									'Parse first line of report
									ParseShow(strResponse, divider, UnitType, ReportType, Line)
								Case 2
									divider = InStr(strResponse, "lcm")
									If divider > 0 Then
										ReportType = "b"
										If InStr(strResponse, "crew") > 0 Then 'planes
											UnitType = "p"
										ElseIf InStr(strResponse, "guns") > 0 Then  'lands
											UnitType = "l"
										ElseIf InStr(strResponse, "rad") > 0 Then  'nukes
											UnitType = "n"
										ElseIf InStr(strResponse, "%") > 0 Then  'sectors
											UnitType = "c"
											ParseShow(strResponse, divider, UnitType, ReportType, Line)
										Else 'ships
											UnitType = "s"
										End If
									Else
										divider = InStr(strResponse, "cargos")
										If divider > 0 Then
											ReportType = "c"
											UnitType = "s"
										Else
											divider = InStr(strResponse, "p  h  x")
											If divider > 0 Then
												ReportType = "s"
												UnitType = "s"
											Else
												divider = InStr(strResponse, "f  a  a")
												If divider > 0 Then
													ReportType = "s"
													UnitType = "l"
												Else
													divider = InStr(strResponse, "acc load")
													If divider > 0 Then
														ReportType = "s"
														UnitType = "p"
													Else
														divider = InStr(strResponse, "blst dam")
														If divider > 0 Then
															ReportType = "s"
															UnitType = "n"
														Else
															divider = InStr(strResponse, "capabilities")
															If divider > 0 Then
																ReportType = "c"
																UnitType = "?"
															Else
																divider = InStr(strResponse, "Bridges require")
																If divider > 0 Then
																	ReportType = "b"
																	UnitType = "b"
																	ParseShow(strResponse, divider, UnitType, ReportType, Line)
																Else
																	divider = InStr(strResponse, "Bridge Towers")
																	If divider > 0 Then
																		ReportType = "b"
																		UnitType = "t"
																		ParseShow(strResponse, divider, UnitType, ReportType, Line)
																	Else
																		divider = InStr(strResponse, "--- level ---")
																		If divider > 0 Then
																			ReportType = "c"
																			UnitType = "c"
																		Else
																			divider = InStr(strResponse, "--  packing")
																			If divider > 0 Then
																				ReportType = "s"
																				UnitType = "c"
																			Else
																				divider = InStr(strResponse, "naviga    packing")
																				If divider > 0 Then
																					ReportType = "s"
																					UnitType = "c"
																				Else
																					divider = InStr(strResponse, "schedule")
																					If divider > 0 Then
																						ReportType = "u"
																						UnitType = "?"
																					Else
																						divider = InStr(strResponse, "item")
																						If divider = 1 Then
																							ReportType = "s"
																							UnitType = "i"
																							ParseShow(strResponse, divider, UnitType, ReportType, Line)
																						Else
																							divider = InStr(strResponse, "  sector type             product  p.e.")
																							If divider = 1 Then
																								ReportType = "skip"
																								UnitType = ""
																								ParseShow(strResponse, divider, UnitType, ReportType, Line)
																							End If
																						End If
																					End If
																				End If
																			End If
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								Case 3
									If UnitType = "c" And ReportType = "c" Then
										divider = InStr(strResponse, "product")
										'save header
										ParseShow(strResponse, divider, UnitType, "h", Line)
									ElseIf UnitType = "c" And ReportType = "s" Then 
										divider = InStr(strResponse, "mcost")
										If divider = 0 Then
											divider = InStr(strResponse, "  0%")
										End If
									ElseIf ReportType = "s" And (UnitType = "l" Or UnitType = "s") Then 
										'skip this line
									ElseIf ReportType = "u" Then 
										'skip this line
									ElseIf UnitType = "?" Then 
										'Try to figure out if we are parsing plane or land capacities
										UnitType = VB.Left(Trim(Mid(strResponse, divider)), 1)
										If UnitType >= "1" And UnitType <= "9" Then
											UnitType = "l"
										Else
											If InStr(Mid(strResponse, divider), "recon") > 0 Then
												UnitType = "l"
											Else
												UnitType = "p"
											End If
										End If
										ParseShow(strResponse, divider, UnitType, ReportType, Line)
									Else
										ParseShow(strResponse, divider, UnitType, ReportType, Line)
									End If
								Case 4
									If ReportType = "s" And (UnitType = "l" Or UnitType = "s") Then
										If UnitType = "l" Then
											divider = InStr(strResponse, "att")
										Else
											divider = InStr(strResponse, "def")
										End If
									ElseIf ReportType = "u" Then 
										'skip this line
									Else
										ParseShow(strResponse, divider, UnitType, ReportType, Line)
									End If
								Case Else
									If ReportType <> "u" Then
										ParseShow(strResponse, divider, UnitType, ReportType, Line)
									End If
							End Select
						End If
						
						'this handles attack and assault commands, with options
					Case READ_STATE_ATTACK
						If nCode = EMP_C_PROMPT Then
							'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
							frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							'PARSE SERVER QUERY
							If Attack_Probe = "y" Then
								Attack_Probe = "s"
							End If
							'Number of mil from city at -2,2 (max 36) :
							If InStr(strResponse, "Number of mil from") > 0 And (Attack_Mil <> -1) Then
								If Attack_Mil = -2 Then
									'figure out the no-support value
									If Attack_DefMil < 0 Then
										'guess defensive mil
										Attack_DefMil = EnemyMil(Attack_Sector)
										If Attack_DefMil > 0 Then
											Attack_DefMil = Attack_DefMil - 5
										ElseIf Attack_DefMil = 0 Then 
											Attack_DefMil = 1
										End If
										Attack_Mil = Int((Attack_DefMil * DefenseStrength(Attack_Sector)) * (1# + (CSng(frmOptions.iDefenseScaling) / 100#)))
										Attack_DefMil = -1
									Else
										Attack_Mil = Int((Attack_DefMil * DefenseStrength(Attack_Sector)) * (1# + (CSng(frmOptions.iDefenseScaling) / 100#)))
									End If
									If Attack_Mil <= 0 Then
										Attack_Mil = 0
									End If
								End If
								n = CShort(Trim(StringBetween(strResponse, "(max", ")")))
								If Attack_Mil >= 0 Then
									If n > Attack_Mil Then
										n = Attack_Mil
										Attack_Mil = 0
									Else
										Attack_Mil = Attack_Mil - n
									End If
								End If
								If Len(Attack_Probe) = 0 Then
									frmDrawMap.lstCmdResult.Items.Add(strResponse & CStr(n))
								End If
								EchoLine(strResponse & CStr(n))
								ServerQuery = True
								SendStringtoServer(CStr(n))
							ElseIf InStr(strResponse, "How many mil to move") > 0 And Attack_Occ >= 0 Then 
								n = CShort(Trim(StringBetween(strResponse, "(", "max)")))
								If n > Attack_Occ Then
									n = Attack_Occ
									Attack_Occ = 0
								Else
									Attack_Occ = Attack_Occ - n
								End If
								EchoLine(strResponse & CStr(n))
								If Len(Attack_Probe) = 0 Then
									frmDrawMap.lstCmdResult.Items.Add(strResponse & CStr(n))
								End If
								ServerQuery = True
								frmDrawMap.lstCmdResult.Items.Add(strResponse & CStr(n))
								SendStringtoServer(CStr(n))
								
								'attack with lar  lt armor #4 in 3,-3 (~ 100%) [ynYNq?]
								'assault with eng  engineer #2 on ls   landing ship (#10) (~ 100%) [ynYNq?]
							ElseIf (InStr(strResponse, "attack with") > 0 Or InStr(strResponse, "assault with") > 0) And Len(Attack_Unit) > 0 Then 
								If LCase(Attack_Unit) = "y" Or LCase(Attack_Unit) = "n" Then
									strTemp = Attack_Unit
								Else
									If InStr(strResponse, "attack with") > 0 Then
										strTemp = StringBetween(strResponse, " #", " in ")
									Else
										strTemp = StringBetween(strResponse, " #", " on ")
									End If
									If InStr("/" & Attack_Unit & "/", "/" & strTemp & "/") > 0 Then
										strTemp = "y"
									Else
										strTemp = "n"
									End If
								End If
								'send response
								If Len(Attack_Probe) = 0 Then
									frmDrawMap.lstCmdResult.Items.Add(strResponse & strTemp)
								End If
								EchoLine(strResponse & strTemp)
								ServerQuery = True
								SendStringtoServer(strTemp)
								
								'Move in with lar  lt armor #4 (~ 100%) [ynYNq?]  y
							ElseIf InStr(strResponse, "Move in with") > 0 And Len(Attack_Occ_Unit) > 0 Then 
								If LCase(Attack_Occ_Unit) = "y" Or LCase(Attack_Occ_Unit) = "n" Then
									strTemp = Attack_Occ_Unit
								Else
									strTemp = StringBetween(strResponse, " #", " (~")
									If InStr("/" & Attack_Occ_Unit & "/", "/" & strTemp & "/") > 0 Then
										strTemp = "y"
									Else
										strTemp = "n"
									End If
								End If
								'send response
								If Len(Attack_Probe) = 0 Then
									frmDrawMap.lstCmdResult.Items.Add(strResponse & strTemp)
								End If
								EchoLine(strResponse & strTemp)
								ServerQuery = True
								SendStringtoServer(strTemp)
							Else
								RespondtoServerQuery(strResponse)
							End If
						ElseIf InStr(strResponse, "blown up by the crew") > 0 Then 
							ParseAndDeleteUnits((strResponse))
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							EchoLine(strResponse)
						Else
							'-1,3 is a 0% Wolf highway with approximately 0 military.
							If InStr(strResponse, "with approximately") > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object ParseAttack(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object VARsec. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								VARsec = ParseAttack(strResponse)
								If IsArray(VARsec) Then
									AddEnemySectInfo(VARsec)
								End If
							End If
							
							If InStr(strResponse, "Final defense strength") > 0 Then
								n = Val(StringBetween(strResponse, ":", vbNullString)) / DefenseStrength(Attack_Sector)
								
								'if amount goes up, something's wrong - must be support
								If Attack_DefMil > 0 And n > Attack_DefMil Then
									Attack_DefMil = -1
								Else
									Attack_DefMil = n
								End If
							End If
							
							If Attack_DefMil > 0 And InStr(strResponse, "Theirs:") > 0 Then
								Attack_DefMil = Attack_DefMil - Val(StringBetween(strResponse, ":", vbNullString))
							End If
							
							If Attack_Probe <> "s" Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							EchoLine(strResponse)
						End If
						
						'This handles parsing the version commands
					Case READ_STATE_VERSION
						If nCode = EMP_C_PROMPT Then
							VersionFirstTimeFlag = 0
							'Update the version info record
							rsVersion.Update()
							'Update the map
							rsVersion.MoveFirst()
							frmDrawMap.SetMapSize(rsVersion.Fields("World X").Value, rsVersion.Fields("World Y").Value)
							If Not DatabaseUpdateinProgress Then
								If Not frmReport.Visible Then ShowReportBox()
							End If
							If Not bXDumpServer And VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
								RequestMetaTables()
							End If
							NextCommand(strResponse)
						Else
							'Add to list boxes
							If Not DatabaseUpdateinProgress Then
								frmReport.AddReportLine(strResponse)
							End If
							EchoLine(strResponse)
							
							'if first time, setup record to be added or modified
							If VersionFirstTimeFlag = 0 Then
								VersionFirstTimeFlag = 99
								If VersionCheck(4, 3, 0) <> WinAceRoutines.enumVersion.VER_LESS Then
									bXDumpServer = True
								Else
									bXDumpServer = False
								End If
								If rsVersion.BOF And rsVersion.EOF Then
									rsVersion.AddNew() 'If empty
								Else
									rsVersion.MoveFirst()
									rsVersion.Edit()
								End If
							End If
							
							'Add to version list
							ParseVersion(strResponse, rsVersion)
						End If
						
						'This handles parsing the nation command
					Case READ_STATE_NATION
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress Then
								If Not frmReport.Visible Then ShowReportBox() 'drk 2/26/03
							End If
							SaveNationVariables()
							NextCommand(strResponse)
						Else
							'Add to list boxes
							If Not DatabaseUpdateinProgress Then
								frmReport.AddReportLine(strResponse)
							End If
							EchoLine(strResponse)
							
							'Add to Nation
							ParseNation(strResponse)
						End If
						
						'This handles parsing the newspaper command 112903 rjk: added
					Case READ_STATE_NEWSPAPER
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress Then
								If bNPHourlyActivity Then
									
									frmReport.ClearReport()
									frmReport.Text = "Hourly Activity Report"
									frmReport.AddReportLine("                                    Hourly Activity Report")
									frmReport.AddReportLine("")
									frmReport.AddReportLine("      Nation           0   1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22  23")
									frmReport.AddReportLine("")
									For NationIndex = 0 To Nations.Count
										'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										strTempLine = PadString(Nations.NationName(NationIndex), 20, False)
										For HourIndex = 0 To 23
											'UPGRADE_WARNING: Couldn't resolve default property of object PadString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											strTempLine = strTempLine + PadString(CStr(ActivityCounts(HourIndex, NationIndex)), 4)
										Next HourIndex
										frmReport.AddReportLine(strTempLine)
									Next NationIndex
								End If
								If (Not frmReport.Visible) And ((Not bNPNoReport) Or bNPHourlyActivity) Then ShowReportBox()
							End If
							If frmRelationsGrid.Visible Then
								frmRelationsGrid.callback()
							End If
							NextCommand(strResponse)
						Else
							'Add to list boxes
							If (Not DatabaseUpdateinProgress) Or bNPNoReport Then
								frmReport.AddReportLine(strResponse)
							End If
							If SubmittedFromCommandLine Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
								EchoLine(strResponse)
							End If
							ParseNewsPaper(strResponse)
						End If
						
						'This handles parsing the player command 120103 rjk: added
					Case READ_STATE_PLAYERS
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress Then
								If Not frmReport.Visible Then ShowReportBox()
							End If
							If frmOptions.playersTimeInterval <> 0 Then
								If frmDrawMap.IsCensusTabActive Then
									frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
								End If
							End If
							NextCommand(strResponse)
						Else
							'Add to list boxes
							If (Not DatabaseUpdateinProgress) Then
								frmReport.AddReportLine(strResponse)
							End If
							If SubmittedFromCommandLine Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
								EchoLine(strResponse)
							End If
							ParsePlayers(strResponse)
						End If
						
						'Shifting the origin requires updating all locations
					Case READ_STATE_ORIGIN
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							n = InStr(strResponse, "Origin at")
							If n > 0 Then
								'shift x,y coordinates in database
								n = n + 9
								n2 = InStr(strResponse, "(")
								If n2 > n Then
									If ParseSectors(sx, sy, Trim(Mid(VB.Left(strResponse, n2 - 1), n))) Then
										ShiftOrigin(sx, sy)
										OriginX = OriginX - sx
										OriginY = OriginY - sy
										'OriginY must be an even number
										If (CShort(OriginY \ 2)) * 2 <> OriginY Then
											OriginY = OriginY + 1
											OriginX = OriginX + 1
											frmDrawMap.DrawHexes()
										End If
									End If
								End If
								EchoLine(strResponse)
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
						End If
						
						'This is for order with auto answer
					Case READ_STATE_ORDER
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							EchoLine(strResponse)
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							SendStringtoServer(vbNullString)
						Else
							EchoLine(strResponse)
						End If
						
						'This is for mission results
					Case READ_STATE_LANDMISSION, READ_STATE_SHIPMISSION, READ_STATE_PLANEMISSION
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							EchoLine(strResponse)
							If Not DatabaseUpdateinProgress Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							Select Case ExpectSpecificResponse
								Case READ_STATE_LANDMISSION
									ParseMissions(strResponse, "l", bDeity)
								Case READ_STATE_SHIPMISSION
									ParseMissions(strResponse, "s", bDeity)
								Case READ_STATE_PLANEMISSION
									ParseMissions(strResponse, "p", bDeity)
							End Select
						End If
						
						'This handles parsing ship orders
					Case READ_STATE_SHOWORDER
						If nCode = EMP_C_PROMPT Then
							If Not DatabaseUpdateinProgress And Not SubmittedFromCommandLine Then
								If Not frmReport.Visible Then ShowReportBox()
							End If
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							'drk 7/14/02. before this was here, we'd lock up the console if we typed "sorder"
							RespondtoServerQuery(strResponse)
						Else
							'Add to list boxes
							If Not DatabaseUpdateinProgress Then
								If Not SubmittedFromCommandLine Then
									frmReport.AddReportLine(strResponse)
								Else
									frmDrawMap.lstCmdResult.Items.Add(strResponse)
								End If
							End If
							EchoLine(strResponse)
							
							'Add to Orders data base
							If InStr(strResponse, "[") > 0 Then
								ParseCargoOrders(strResponse, bDeity)
							Else
								ParseShipOrders(strResponse, bDeity)
							End If
						End If
						
					Case READ_STATE_EDIT
						If nCode = EMP_C_PROMPT Then
							strReply1 = vbNullString
							strReply2 = vbNullString
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							If InStr(strResponse, "xxxx") > 0 And Len(strReply1) > 0 Then
								n = InStr(strReply1, ":")
								If n = 0 Then
									strReply2 = strReply1
									strReply1 = ":"
								ElseIf n = 1 Then 
									strReply2 = vbNullString
									strReply1 = vbNullString
								Else
									strReply2 = VB.Left(strReply1, n - 1)
									strReply1 = Mid(strReply1, n + 1)
								End If
								ServerQuery = True
								EchoLine(strResponse & strReply2)
								SendStringtoServer(strReply2)
								strReply2 = vbNullString
							Else
								RespondtoServerQuery(strResponse)
							End If
						Else
							'121403 rjk: Printed the edit output to main commmand as well
							If SubmittedFromCommandLine Then
								frmDrawMap.lstCmdResult.Items.Add(strResponse)
							End If
							EchoLine(strResponse)
						End If
						
						'The remaining special cases are used to control the empire logon technique
					Case READ_STATE_FIRSTSERVERINIT, READ_STATE_USERCMDOK, READ_STATE_COUNCMDOK, READ_STATE_PASSCMDOK, READ_STATE_KILLPROC, READ_STATE_PLAYINIT, READ_STATE_LISTSANC, READ_STATE_OPTIONS
						ProcessServerResponse_ExpectSpecificResponse_Logon(strResponse, ExpectSpecificResponse, nCode, strData)
						
						
					Case Else
						
						If nCode = EMP_C_PROMPT Then
							NextCommand(strResponse)
						ElseIf nCode = EMP_C_FLUSH Then 
							RespondtoServerQuery(strResponse)
						Else
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
							EchoLine(strResponse)
						End If
						
				End Select
				
				
				'this ends the special cases
				
			Else
				Select Case nCode
					Case EMP_C_PROMPT 'Current Command has finished
						If BatchFileinProgress Then
							'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
							If frmDrawMap.lstCmdResult.NewIndex > 0 Then
								'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
								frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
							End If
							NextCommand(strResponse)
						Else
							NextCommand(strResponse)
						End If
						
					Case EMP_C_DATA
						'Put response into list box
						EchoLine(strResponse)
						If DisplayCmd Then
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
						End If
						NewLineCount = NewLineCount + 1
					Case EMP_C_HUB_DATA
						'Put response into list box
						If InStr(strResponse, "DUMP SECTOR") > 0 Then
							Line = 2
							ExpectSpecificResponse = READ_STATE_SECTOR_DUMP
						ElseIf InStr(strResponse, "DUMP SHIP") > 0 Then 
							Line = 1
							ExpectSpecificResponse = READ_STATE_SHIP_DUMP
							ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_SHIP_DUMP, nCode, Line, rsShip, strResponse, "ship", HaveShips, tsShip)
						ElseIf InStr(strResponse, "DUMP PLANE") > 0 Then 
							Line = 1
							ExpectSpecificResponse = READ_STATE_PLANE_DUMP
							ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_PLANE_DUMP, nCode, Line, rsPlane, strResponse, "plane", HavePlanes, tsPlane)
						ElseIf InStr(strResponse, "DUMP LAND") > 0 Then 
							Line = 1
							ExpectSpecificResponse = READ_STATE_LAND_DUMP
							ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(READ_STATE_LAND_DUMP, nCode, Line, rsLand, strResponse, "unit", HaveLands, tsLand)
						ElseIf InStr(strResponse, "DUMP NUKE") > 0 Then 
							Line = 2
							ExpectSpecificResponse = READ_STATE_NUKE_DUMP
						ElseIf InStr(strResponse, "DUMP LOST ITEMS") > 0 Then 
							Line = 2
							rsSectors.Index = "PrimaryKey"
							rsShip.Index = "PrimaryKey"
							rsPlane.Index = "PrimaryKey"
							rsLand.Index = "PrimaryKey"
							rsNuke.Index = "PrimaryKey"
							ExpectSpecificResponse = READ_STATE_LOST_DUMP
						ElseIf InStr(strResponse, "XDUMP") > 0 Then 
							Line = 0
							ExpectSpecificResponse = READ_STATE_XDUMP
							ParseXdump(strResponse, UpdateTimeStamp)
						End If
						EchoLine(strResponse)
						If DisplayCmd Then
							frmDrawMap.lstCmdResult.Items.Add(strResponse)
						End If
						NewLineCount = NewLineCount + 1
					Case EMP_C_HUB_COMMAND
						EchoLine("HUB Command:" & strResponse)
					Case EMP_C_HUB_PROMPT
						EchoLine("HUB Prompt:" & strResponse)
					Case EMP_C_FLUSH
						RespondtoServerQuery(strResponse)
					Case EMP_C_EXIT 'Current Command has finished
						MsgBox("Empire Server has broken Connection")
					Case Else
						MsgBox("Unexpected Code " & CStr(nCode) & strResponse)
				End Select
			End If
			
			'Scroll the list box every 6 lines
			If NewLineCount > 5 Then
				'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				If List1.NewIndex < -1 Or List1.NewIndex > 32000 Then
					List1.Items.Clear()
				Else
					'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
					List1.TopIndex = List1.NewIndex
				End If
				NewLineCount = 0
			End If
		End While
		
		'Scroll the list box at the end
		'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		If List1.NewIndex < -1 Or List1.NewIndex > 32000 Then
			List1.Items.Clear()
		Else
			'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				'UPGRADE_WARNING: Untranslated statement in (Declarations). Please check source code.
		End If
		
		'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		If DisplayCmd And frmDrawMap.lstCmdResult.NewIndex <> -1 Then '120203 rjk: Added check for when no command was added.
			'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
			If frmDrawMap.lstCmdResult.NewIndex < -1 Or frmDrawMap.lstCmdResult.NewIndex > 32000 Then
				frmDrawMap.lstCmdResult.Items.Clear()
			Else
				'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
			End If
		End If
		
		NewLineCount = 0
		
		Exit Sub
Empire_Error: 
		EmpireError("ProcessServerResponse", CStr(nCode) & ":" & CStr(ExpectSpecificResponse), strBuffer)
	End Sub
	
	'This section is for when a Plane/land/ship dump has been requested
	Private Sub ProcessServerResponse_ExpectSpecificResponse_PLS_Dump(ByRef ExpectSpecificResponse As Short, ByRef nCode As Short, ByRef Line As Short, ByRef rsUnits As DAO.Recordset, ByRef strResponse As String, ByRef key As String, ByRef HaveUnits As Boolean, ByRef tsUnit As Integer)
		Dim n As Short
		Static bFillGrid As Boolean '102703 rjk: Added to detect when grid is required.
		Static strFields() As String
		
		If nCode = EMP_C_PROMPT Then
			If ExpectSpecificResponse = READ_STATE_SHIP_DUMP Then CarrierDumpRequested = False
			NextCommand(strResponse)
			Line = 0
		ElseIf nCode = EMP_C_HUB_COMMAND Then 
			NextCommand()
			Line = 0
		ElseIf nCode = EMP_C_FLUSH Then 
			ExecuteEmpireCmd("*")
		Else
			Line = Line + 1
			If Line = 1 Then
				rsUnits.Index = "PrimaryKey"
				bFillGrid = False
			ElseIf Line > 3 Then 
				n = InStr(strResponse, key)
				If n > 0 Then
					'We are done - this is plane/ship/land count
					
					If rsUnits.BOF And rsUnits.EOF Then
						HaveUnits = False
					Else
						HaveUnits = True
					End If
					Line = 0
					EchoLine(strResponse)
					ExpectSpecificResponse = 0
					If bFillGrid Then '102703 rjk: Only update grid when required.
						frmDrawMap.FillGrid()
					End If
					frmDrawMap.Refresh()
				Else
					If SubmittedFromCommandLine Then
						EchoLine(strResponse)
					End If
					If ProcessDump(strResponse, rsUnits, strFields) Then '102703 rjk: Added to detect when grid is required.
						bFillGrid = True
					End If
				End If
			ElseIf Line = 2 And UpdateTimeStamp Then 
				If InStr(strResponse, "DUMP") > 0 Then
					strResponse = Trim(strResponse)
					n = InStr(strResponse, " ")
					n = InStr(n + 1, strResponse, " ")
					'101403 rjk: Got out of step and cause run-time errors processing PLANE dump when expecting LAND dump
					'            Changed to reduce the occurance of run-time errors
					'If ExpectSpecificResponse = READ_STATE_LAND_DUMP Then n = InStr(n + 1, strResponse, " ")
					If InStr(strResponse, "DUMP LAND UNITS") Then n = InStr(n + 1, strResponse, " ")
					tsUnit = CInt(Trim(Mid(strResponse, n + 1))) + DUMP_DELAY
				End If
				If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
					EchoLine(strResponse)
				End If
			Else
				If frmOptions.bDisplayDumpHeaders And SubmittedFromCommandLine Then
					EchoLine(strResponse)
				End If
				If Line = 3 Then
					strFields = Split(strResponse) '020304 rjk: Added to process the dump based on the header information
				End If
			End If
		End If
		
	End Sub
	
	Private Sub ProcessServerResponse_ExpectSpecificResponse_Logon(ByRef strResponse As String, ByRef ExpectSpecificResponse As Short, ByRef nCode As Short, ByRef strData As String)
		'Expecting Empire server Ready - Send user name
		'022004 rjk: Removed extra space commands, test server does not work them.
		Dim nClientProto As Short
		Select Case ExpectSpecificResponse
			Case READ_STATE_FIRSTSERVERINIT
				'
				EchoLine(strResponse)
				frmDrawMap.lstCmdResult.Items.Add(strResponse)
				If nCode = EMP_C_INIT Then
					If bListSanc Then
						ExpectSpecificResponse = READ_STATE_LISTSANC
						WriteStringToSocket("sanc" & vbCrLf)
						frmReport.Text = "List Sanctuaries"
						frmReport.ClearReport()
					ElseIf InStr(strResponse, "Empire Hub Server") > 0 Then 
						bEmpireHub = True
						tokenStatus = enumToken.tsIdle
						If Not bUTF8 Then
							ExpectSpecificResponse = READ_STATE_PLAYINIT
							WriteStringToSocket("play " & Chr(34) & LoginUser & Chr(34) & " " & Chr(34) & CountryName & Chr(34) & " " & Chr(34) & Password & Chr(34) & vbCrLf)
							If Me.bEmpireHub Then
								frmChat.Show()
							End If
						End If
					ElseIf bUTF8 Then 
						ExpectSpecificResponse = READ_STATE_OPTIONS
						WriteStringToSocket("options utf-8" & vbCrLf)
					Else
						ExpectSpecificResponse = READ_STATE_USERCMDOK
						WriteStringToSocket("user " & Chr(34) & LoginUser & Chr(34) & vbCrLf)
					End If
				Else
					MsgBox("Received Wrong Code" & vbCrLf & "Expected: 2 Empire server ready")
					FailConnect()
				End If
				
			Case READ_STATE_OPTIONS
				If nCode = EMP_C_DATA Then
					Exit Sub
				End If
				If nCode <> EMP_C_CMDOK Then
					MsgBox("UTF-8 Option is not available on this server", MsgBoxStyle.OKOnly)
					bUTF8 = False
				End If
				If bEmpireHub Then
					ExpectSpecificResponse = READ_STATE_PLAYINIT
					WriteStringToSocket("play " & Chr(34) & LoginUser & Chr(34) & " " & Chr(34) & CountryName & Chr(34) & " " & Chr(34) & Password & Chr(34) & vbCrLf)
					If Me.bEmpireHub Then
						frmChat.Show()
					End If
				Else
					ExpectSpecificResponse = READ_STATE_USERCMDOK
					WriteStringToSocket("user " & Chr(34) & LoginUser & Chr(34) & vbCrLf)
				End If
				
				'Expecting Hello, so we can send country code
			Case READ_STATE_USERCMDOK
				
				EchoLine(strResponse)
				frmDrawMap.lstCmdResult.Items.Add(strResponse)
				If nCode = EMP_C_CMDOK Then
					ExpectSpecificResponse = READ_STATE_COUNCMDOK
					WriteStringToSocket("coun " & Chr(34) & CountryName & Chr(34) & vbCrLf)
				Else
					'FailConnect(2);
					MsgBox("Received Wrong Code" & vbCrLf & "Expected: 0 hello USER")
					FailConnect()
				End If
				
				'Expected Password Request
			Case READ_STATE_COUNCMDOK
				
				EchoLine(strResponse)
				frmDrawMap.lstCmdResult.Items.Add(strResponse)
				If nCode = EMP_C_CMDOK Then
					ExpectSpecificResponse = READ_STATE_PASSCMDOK
					WriteStringToSocket("pass " & Chr(34) & Password & Chr(34) & vbCrLf)
				Else
					'FailConnect(3);
					MsgBox("No Such Country" & vbCrLf & "Expected: 0 country name COUN",  , "Login Error")
					FailConnect()
				End If
				
				'Password Confirm is Ok
			Case READ_STATE_PASSCMDOK
				
				EchoLine(strResponse)
				frmDrawMap.lstCmdResult.Items.Add(strResponse)
				If nCode = EMP_C_CMDOK Then
					
					If bKillSession Then
						'This is the Kill Routine to Kill Hung Session
						ExpectSpecificResponse = READ_STATE_KILLPROC
						WriteStringToSocket("kill" & vbCrLf)
					Else
						ExpectSpecificResponse = READ_STATE_PLAYINIT
						WriteStringToSocket("play" & vbCrLf)
					End If
				Else
					MsgBox("Bad Password" & vbCrLf & "Expected: 0 password")
					FailConnect()
				End If
				
				'Processed Kill Request
			Case READ_STATE_KILLPROC
				EchoLine("Empire Server returned:")
				EchoLine(strResponse)
				'220104 rjk: Added to save going back to login screen and changing check box
				If strResponse = "closed socket of offending job" Or strResponse = "country not in use" Then
					ExpectSpecificResponse = READ_STATE_PLAYINIT
					WriteStringToSocket("play" & vbCrLf)
				Else
					FailConnect()
				End If
				
				'Ready to go !
			Case READ_STATE_PLAYINIT
				
				If nCode = EMP_C_INIT Then
					ExpectSpecificResponse = READ_STATE_REPORT
					frmReport.Text = GameName & " Greetings"
					CmdinProgress = False
					UpdateLights()
					frmDrawMap.lstCmdResult.Items.Add("Sucessful Logon")
					frmDrawMap.lstCmdResult.Items.Add("Updating Database...")
					
					'see if next arg matches this CLIENTPROTO
					nClientProto = Asc(Mid(strData, 3, 1)) - Asc("0")
					If (nClientProto <> EMP_CLIENT_PROTO) Then
						MsgBox("Warning: Got Client Proto =" & CStr(nClientProto) & "; expected " & CStr(EMP_CLIENT_PROTO))
					End If
					'Add Successful connection Code here
				Else
					EchoLine(strResponse)
					MsgBox("Login Error: Received Wrong Code" & vbCrLf & "Expected: 2 CLIENTPROTO" & vbCrLf & "Got: " & strResponse)
					FailConnect()
				End If
				
				'List out countries still in sanctuary
			Case READ_STATE_LISTSANC
				If nCode = EMP_C_DATA Then
					frmReport.AddReportLine(strResponse)
					'250304 rjk: Removed the check for EMP_C_CMDOK so the user sees the output
					'            and then can continue to select a new country
					EchoLine(strResponse)
				ElseIf nCode = EMP_C_CMDOK Then 
					frmReport.ShowDialog()
					frmEmpireLogin.chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked
					FailConnect()
				Else
					frmReport.AddReportLine(strResponse)
					EchoLine(strResponse)
					frmReport.ShowDialog()
					frmEmpireLogin.chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked
					FailConnect()
				End If
		End Select
	End Sub
	
	Public Sub NextCommand(Optional ByRef strResponse As Object = Nothing)
		On Error GoTo Empire_Error
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(strResponse) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strResponse = vbNullString
		End If
		
		'Add a blank line to seperate commands, suppress internal commands.
		If Len(strResponse) > 0 Then
			If DisplayCmd Or Not frmOptions.bSuppressUpdateCommands Then
				EchoLine(vbNullString)
			End If
		End If
		'UPGRADE_ISSUE: ListBox property List1.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		List1.TopIndex = List1.NewIndex
		
		CmdinProgress = False
		UpdateLights()
		ExpectSpecificResponse = 0
		DisplayCmd = True
		UpdateTimeStamp = False
		
		'Check for BTU's left and display in the status box
		Dim p1 As Short
		If Len(strResponse) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			p1 = InStr(strResponse, " ")
			If p1 > 0 And p1 < Len(strResponse) Then
				'frmDrawMap.lblBTU = "BTU: " + mid$(strResponse, p1 + 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object strResponse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmDrawMap.sb1.Items.Item("BTU").Text = " BTU: " & Mid(strResponse, p1 + 1) & " "
			End If
		End If
		
		'if there are more commands, execute the next one
		If CmdQ.CmdsPending > 0 Then
			TransferEmpireCmd()
		Else
			If bEmpireHub Then
				Select Case tokenStatus
					Case enumToken.tsReceived
						SendStringtoServer("ReLeAsE")
						CmdinProgress = False
						UpdateLights()
						tokenStatus = enumToken.tsIdle
					Case enumToken.tsIdle
						'nothing to do
					Case Else
						Debug.Print("NextCommand: error invalid state for Token")
				End Select
			End If
		End If
		Exit Sub
Empire_Error: 
		EmpireError("NextCommand", vbNullString, vbNullString)
	End Sub
	
	Private Sub TransferEmpireCmd()
		If CmdQ.CmdsPending > 0 Then
			If bEmpireHub And Not RecordingScriptFile Then
				Select Case tokenStatus
					Case enumToken.tsIdle
						SendStringtoServer("ToKeN")
						tokenStatus = enumToken.tsSent
					Case enumToken.tsReceived
						ExecuteEmpireCmd(CmdQ.GetCommand)
					Case enumToken.tsSent
						Debug.Print("TransferEmpireCmd: command finish waiting for token")
					Case Else
						Debug.Print("TransferEmpireCmd: error invalid state for Token")
						Exit Sub
				End Select
			Else
				ExecuteEmpireCmd(CmdQ.GetCommand)
			End If
		End If
	End Sub
	
	Private Sub RespondtoServerQuery(ByRef strResponse As String)
		On Error GoTo Empire_Error
		
		If CmdinProgress Then
			ServerQuery = True
		End If
		
		Dim strTemp As String
		If BatchFileinProgress Then
			'if there is a response pending, use it
			If CmdQ.QueryResponseNext Then
				strTemp = Mid(CmdQ.GetCommand, 2)
				SendStringtoServer(strTemp)
				EchoLine(strResponse & " " & strTemp)
				frmDrawMap.lstCmdResult.Items.Add(strResponse & " " & strTemp)
				Exit Sub
			End If
		End If
		
		'Put repsonse into list box
		EchoLine(strResponse)
		NewLineCount = NewLineCount + 1
		
		'Add to Screen if necessesary
		' Dim j As Integer   removed efj 8/2003
		' Dim i As Integer   removed efj 8/2003
		If DisplayCmd Then
			frmDrawMap.lstCmdResult.Items.Add(strResponse)
		Else
			'112903 rjk: Not sure what Daniel was working on.
			'If foo = 0 Then
			'    MsgBox "WinACE bug! Could you please e-mail drk@unxsoft.com and tell him exactly what you were doing when this box appeared.  Thank you.", vbCritical + vbOKOnly, "Bug!"
			'    foo = 1 'if this is something that happens a lot, only annoy the user once
			'End If
			'fixme!! 2/27/02 showstopper
			'j = frmReport.lstReport.ListCount
			'For i = 0 To j - 1
			'    frmDrawMap.lstCmdResult.AddItem frmReport.lstReport.List(i)
			'Next i
			frmDrawMap.lstCmdResult.Items.Add(strResponse)
			DisplayCmd = True
		End If
		
		'Scroll the list box at the end
		'UPGRADE_ISSUE: ListBox property lstCmdResult.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
		frmDrawMap.lstCmdResult.TopIndex = frmDrawMap.lstCmdResult.NewIndex
		
		'put the focus in the response box
		On Error Resume Next
		If Not SubmitedfromCmdWindow And CmdinProgress Then
			frmDrawMap.txtEntry.Focus()
		End If
		
		Exit Sub
Empire_Error: 
		EmpireError("RespondtoServerQuery", vbNullString, strResponse)
		
	End Sub
	
	Private Sub ShowReportBox()
		On Error GoTo Empire_Error
		
		If SubmittedFromCommandLine Then
			Exit Sub
		End If
		
		'Minimize other boxes
		
		'if land is Not minimized, then minimize it
		If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized Then
			Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
		End If
		
		'Show Report
		frmReport.Show()
		Exit Sub
Empire_Error: 
		EmpireError("ShowReportBox", vbNullString, vbNullString)
	End Sub
	
	Private Sub FailConnect()
		On Error GoTo Empire_Error
		
		'try to reconnect, or change login info
		Me.ConnectToServer()
		
		'If still connected, then resubmit signon
		If Me.ConnectedtoHost Then
			If bListSanc Then
				ExpectSpecificResponse = READ_STATE_LISTSANC
				WriteStringToSocket("sanc" & vbCrLf)
			Else
				ExpectSpecificResponse = READ_STATE_USERCMDOK
				WriteStringToSocket("user " & Chr(34) & LoginUser & Chr(34) & vbCrLf)
			End If
		Else
			ExpectSpecificResponse = 0
			CmdinProgress = False
			UpdateLights()
		End If
		Exit Sub
Empire_Error: 
		EmpireError("FailConnect", vbNullString, vbNullString)
		
	End Sub
	
	'This routine routes output to the command window or to a file
	Private Sub EchoLine(ByRef strLine As Object)
		If OutputToFile Then
			On Error GoTo Echo_Error
			PrintLine(OutputFileNumber, strLine)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object strLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			List1.Items.Add(strLine)
			PrintLine(LogFileNumber, strLine)
		End If
		Exit Sub
		
Echo_Error: 
		OutputToFile = False
	End Sub
End Class