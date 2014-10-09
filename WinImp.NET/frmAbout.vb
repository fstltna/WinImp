Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAbout
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'100103 rjk: Added to check for nukes during startup
	'100303 rjk: Changed the wolfpack version 4.2.12
	'102503 rjk: Changed the support to Sourceforge.net site.
	'270104 rjk: Ensure the delivery colors are set for new install
	'180404 rjk: Added Esc code for Cancel button, make OK the default, changed Apply to OK
	'190404 rjk: Changed the wolfpack version 4.2.14
	'210404 rjk: Added Initialization for tz (TimeZone)
	'230504 rjk: Changed the wolfpack version 4.2.15
	'300505 rjk: Changed the wolfpack version 4.2.21 and fixed spelling mistake.
	
	' Reg Key Security Options...
	Const READ_CONTROL As Integer = &H20000
	Const KEY_QUERY_VALUE As Integer = &H1
	Const KEY_SET_VALUE As Integer = &H2
	Const KEY_CREATE_SUB_KEY As Integer = &H4
	Const KEY_ENUMERATE_SUB_KEYS As Integer = &H8
	Const KEY_NOTIFY As Integer = &H10
	Const KEY_CREATE_LINK As Integer = &H20
	Const KEY_ALL_ACCESS As Double = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
	
	' Reg Key ROOT Types...
	Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Const ERROR_SUCCESS As Short = 0
	Const REG_SZ As Short = 1 ' Unicode nul terminated string
	Const REG_DWORD As Short = 4 ' 32-bit number
	
	Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
	Const gREGVALSYSINFOLOC As String = "MSINFO"
	Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
	Const gREGVALSYSINFO As String = "PATH"
	
	Private Declare Function RegOpenKeyEx Lib "advapi32"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Private Declare Function RegQueryValueEx Lib "advapi32"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer
	
	
	Private Sub cmdSysInfo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSysInfo.Click
		Call StartSysInfo()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Text = "About " & My.Application.Info.Title
		lblVersion.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
		lblTitle.Text = My.Application.Info.Title & " - " & My.Application.Info.Description
		'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
		Me.SetBounds(0, 0, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		GetTimeZoneInformation(tz)
		'If draw map is already up, enable buttons, otherwise turn on timer
		'to run as splash screen.
		Dim frm As System.Windows.Forms.Form
		Dim bSplash As Boolean
		bSplash = True
		For	Each frm In My.Application.OpenForms
			If frm.Name = "frmDrawMap" Then
				bSplash = False
			End If
		Next frm
		
		If Not bSplash Then
			Me.cmdOK.Visible = True
			Me.cmdSysInfo.Visible = True
		Else
			Timer1.Enabled = True
			Timer1.Interval = 50
		End If
	End Sub
	
	Public Sub StartSysInfo()
		On Error GoTo SysInfoErr
		
		' Dim rc As Long    8/2003 efj  removed
		Dim SysInfoPath As String
		
		' Try To Get System Info Program Path\Name From Registry...
		If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
			' Try To Get System Info Program Path Only From Registry...
		ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then 
			' Validate Existance Of Known 32 Bit File Version
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If (Dir(SysInfoPath & "\MSINFO32.EXE") <> vbNullString) Then
				SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
				
				' Error - File Can Not Be Found...
			Else
				GoTo SysInfoErr
			End If
			' Error - Registry Entry Can Not Be Found...
		Else
			GoTo SysInfoErr
		End If
		
		Call Shell(SysInfoPath, AppWinStyle.NormalFocus)
		
		Exit Sub
SysInfoErr: 
		MsgBox("System Information Is Unavailable At This Time", MsgBoxStyle.OKOnly)
	End Sub
	
	Public Function GetKeyValue(ByRef KeyRoot As Integer, ByRef KeyName As String, ByRef SubKeyRef As String, ByRef KeyVal As String) As Boolean
		Dim i As Integer ' Loop Counter
		Dim rc As Integer ' Return Code
		Dim hKey As Integer ' Handle To An Open Registry Key
		'Dim hDepth As Long                    8/2003 efj  removed
		Dim KeyValType As Integer ' Data Type Of A Registry Key
		Dim tmpVal As String ' Tempory Storage For A Registry Key Value
		Dim KeyValSize As Integer ' Size Of Registry Key Variable
		'------------------------------------------------------------
		' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
		'------------------------------------------------------------
		rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Error...
		
		tmpVal = New String(Chr(0), 1024) ' Allocate Variable Space
		KeyValSize = 1024 ' Mark Variable Size
		
		'------------------------------------------------------------
		' Retrieve Registry Key Value...
		'------------------------------------------------------------
		rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize) ' Get/Create Key Value
		
		If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Errors
		
		If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then ' Win95 Adds Null Terminated String...
			tmpVal = VB.Left(tmpVal, KeyValSize - 1) ' Null Found, Extract From String
		Else ' WinNT Does NOT Null Terminate String...
			tmpVal = VB.Left(tmpVal, KeyValSize) ' Null Not Found, Extract String Only
		End If
		'------------------------------------------------------------
		' Determine Key Value Type For Conversion...
		'------------------------------------------------------------
		Select Case KeyValType ' Search Data Types...
			Case REG_SZ ' String Registry Key Data Type
				KeyVal = tmpVal ' Copy String Value
			Case REG_DWORD ' Double Word Registry Key Data Type
				For i = Len(tmpVal) To 1 Step -1 ' Convert Each Bit
					KeyVal = KeyVal & Hex(Asc(Mid(tmpVal, i, 1))) ' Build Value Char. By Char.
				Next 
				KeyVal = VB6.Format("&h" & KeyVal) ' Convert Double Word To String
		End Select
		
		GetKeyValue = True ' Return Success
		rc = RegCloseKey(hKey) ' Close Registry Key
		Exit Function ' Exit
		
GetKeyError: ' Cleanup After An Error Has Occured...
		KeyVal = vbNullString ' Set Return Val To Empty String
		GetKeyValue = False ' Return Failure
		rc = RegCloseKey(hKey) ' Close Registry Key
	End Function
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		' Static ShowMap As Boolean
		Dim n As Short
		
		'make sure there is a default map size.  If left to
		'zero, it will cause errors.
		MapSizeX = 48
		MapSizeY = 48
		
		'drk@unxsoft.com 10/27/02. I was getting tired of getting an error message
		'when I simply pressed cancel on the frmEmpireLogin
		Do 
			'Show the select game and connection form
			frmGameChoice.ShowDialog()
			frmGameChoice.Close()
			Me.Refresh()
			
			'If open was Successfull, go on
			If Not IsDBOpen Then
				MsgBox("Error opening Database - Program will close", MsgBoxStyle.OKOnly)
				End
			End If
			Me.Refresh()
			
		Loop Until frmEmpCmd.ConnectedtoHost Or frmEmpireLogin.bOffline '092603 rjk: Added for Offline mode
		
		'Check that there is a server connection
		frmEmpCmd.CheckLostConnection()
		Me.Refresh()
		
		'Start up empire command interface
		frmEmpCmd.WindowState = System.Windows.Forms.FormWindowState.Minimized
		frmEmpCmd.Show()
		frmTelegram.WindowState = System.Windows.Forms.FormWindowState.Minimized
		frmTelegram.Show()
		
		'force to check for units the first time
		frmEmpCmd.HaveLands = True
		frmEmpCmd.HavePlanes = True
		frmEmpCmd.HaveShips = True
		frmEmpCmd.HaveNukes = True '100103 rjk: Added to check for nukes during startup
		
		'Startup main map
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmDrawMap)
		
		'load default colors
		SetDefaultBasicColors()
		SetDefaultGradColors()
		SetDefaultComparisionColors()
		SetDefaultPlayerColors()
		SetDefaultDeliveryColors() '270104 rjk: Ensure the delivery colors are set for new install
		GradColorDspVal = True
		
		'load saved colors from registry
		Dim strColor As String
		For n = 1 To NUMBER_OF_BASIC_COLORS
			strColor = GetSetting(APPNAME, "Colors", "BasicColor" & CStr(n), vbNullString)
			If Len(strColor) > 0 And IsNumeric(strColor) Then
				BasicColors(n) = CInt(strColor)
			End If
			strColor = GetSetting(APPNAME, "Colors", "BasicText" & CStr(n), vbNullString)
			If Len(strColor) > 0 And IsNumeric(strColor) Then
				BasicText(n) = CInt(strColor)
			End If
		Next n
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			strColor = GetSetting(APPNAME, "Colors", "PlayerColors" & CStr(n), vbNullString)
			If Len(strColor) > 0 And IsNumeric(strColor) Then
				PlayerColors(n) = CInt(strColor)
			End If
			strColor = GetSetting(APPNAME, "Colors", "PlayerText" & CStr(n), vbNullString)
			If Len(strColor) > 0 And IsNumeric(strColor) Then
				PlayerText(n) = CInt(strColor)
			End If
		Next n
		
		'load the nation list
		If Not (rsNations.BOF And rsNations.EOF) Then
			rsNations.MoveFirst()
			Nations.Clear()
			While Not rsNations.EOF
				'drk@unxsoft.com : see the note in frmEmpCmd as to why I truncated the next line.
				Nations.AddNation(rsNations.Fields("Number").Value, rsNations.Fields("Name").Value) ', rsNations.Fields("yours"), rsNations.Fields("theirs")
				If rsNations.Fields("Name").Value = frmEmpCmd.CountryName Then
					CountryNumber = rsNations.Fields("Number").Value
				End If
				rsNations.MoveNext()
			End While
		End If
		
		Timer1.Enabled = False
		
		'update lights on draw map
		UpdateLights()
		
		'Put up the splash screen
		On Error GoTo Timer1_Exit
		frmDrawMap.Show()
		frmDrawMap.UpdateTimer.Interval = 1000
		
		If Not frmDrawMap.MapValid Then
			'UPGRADE_ISSUE: PictureBox property picMap.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			frmDrawMap.picMap.AutoRedraw = True
			frmDrawMap.picMap.Image = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\WinACE.gif")
			frmDrawMap.picMap.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(frmDrawMap.vsMap.Left) - VB6.PixelsToTwipsX(frmDrawMap.picMap(1).Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(frmDrawMap.hsMap.Top) - VB6.PixelsToTwipsY(frmDrawMap.picMap(1).Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
			frmDrawMap.picMap.Visible = True
		End If
		
		
		'If the report box is already up, need to bring it to the for-front
		Dim frm As System.Windows.Forms.Form
		For	Each frm In My.Application.OpenForms
			If frm.Name = "frmReport" Then
				On Error Resume Next 'Check for error in case box is just being created
				frmReport.Activate()
				On Error GoTo 0
			End If
		Next frm
		
Timer1_Exit: 
		Me.Close()
		
	End Sub
End Class