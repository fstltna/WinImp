VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5424
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   8292
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3737.529
   ScaleMode       =   0  'User
   ScaleWidth      =   7789.433
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   3960
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   3684
      ScaleWidth      =   2844
      TabIndex        =   10
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acknowledgements"
      Height          =   1695
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
      Begin VB.Label lblAck 
         Caption         =   "Updates and maintence since v2.3.0: Ron Koenderink (RimFire)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblAck 
         Caption         =   "Updates and maintence since v2.2.0: Daniel Kiracofe (Bwian)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblAck 
         Caption         =   "Much thanks to Jon Butler for the original WS Empire code which provided a good server connection template."
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label lblAck 
         Caption         =   "Collaboration and Testing:  Ken McDonald (also Escher)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label lblAck 
         Caption         =   "Original Design and Programming:  Jim Simons (Escher)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6960
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   6960
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Questions, comments, and bug reports should be submitted to WinACE group at sourceforge.net/projects/winace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      Width           =   5085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.727
      X2              =   7549.888
      Y1              =   2898.928
      Y2              =   2898.928
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":2CB5
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   5085
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   4845
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   4725
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":2D98
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   6735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "About " & App.title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblTitle.Caption = App.title + " - " + App.Comments
'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Me.Move 0, 0
    
GetTimeZoneInformation tz
'If draw map is already up, enable buttons, otherwise turn on timer
'to run as splash screen.
Dim frm As Form
Dim bSplash As Boolean
bSplash = True
For Each frm In Forms
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
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    'Dim hDepth As Long                    8/2003 efj  removed
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = vbNullString                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Timer1_Timer()
' Static ShowMap As Boolean
Dim n As Integer

'make sure there is a default map size.  If left to
'zero, it will cause errors.
MapSizeX = 48
MapSizeY = 48

'drk@unxsoft.com 10/27/02. I was getting tired of getting an error message
'when I simply pressed cancel on the frmEmpireLogin
Do
    'Show the select game and connection form
    frmGameChoice.Show vbModal
    Unload frmGameChoice
    Me.Refresh

    'If open was Successfull, go on
    If Not IsDBOpen Then
        MsgBox "Error opening Database - Program will close", vbOKOnly
        End
    End If
    Me.Refresh
    
Loop Until frmEmpCmd.ConnectedtoHost Or frmEmpireLogin.bOffline '092603 rjk: Added for Offline mode

'Check that there is a server connection
frmEmpCmd.CheckLostConnection
Me.Refresh

'Start up empire command interface
frmEmpCmd.WindowState = vbMinimized
frmEmpCmd.Show vbModeless
frmTelegram.WindowState = vbMinimized
frmTelegram.Show vbModeless

'force to check for units the first time
frmEmpCmd.HaveLands = True
frmEmpCmd.HavePlanes = True
frmEmpCmd.HaveShips = True
frmEmpCmd.HaveNukes = True '100103 rjk: Added to check for nukes during startup

'Startup main map
Load frmDrawMap

'load default colors
SetDefaultBasicColors
SetDefaultGradColors
SetDefaultComparisionColors
SetDefaultPlayerColors
SetDefaultDeliveryColors '270104 rjk: Ensure the delivery colors are set for new install
GradColorDspVal = True

'load saved colors from registry
Dim strColor As String
For n = 1 To NUMBER_OF_BASIC_COLORS
    strColor = GetSetting(APPNAME, "Colors", "BasicColor" + CStr(n), vbNullString)
    If Len(strColor) > 0 And IsNumeric(strColor) Then
        BasicColors(n) = CLng(strColor)
    End If
    strColor = GetSetting(APPNAME, "Colors", "BasicText" + CStr(n), vbNullString)
    If Len(strColor) > 0 And IsNumeric(strColor) Then
        BasicText(n) = CLng(strColor)
    End If
Next n
For n = 1 To NUMBER_OF_PLAYER_COLORS
    strColor = GetSetting(APPNAME, "Colors", "PlayerColors" + CStr(n), vbNullString)
    If Len(strColor) > 0 And IsNumeric(strColor) Then
        PlayerColors(n) = CLng(strColor)
    End If
    strColor = GetSetting(APPNAME, "Colors", "PlayerText" + CStr(n), vbNullString)
    If Len(strColor) > 0 And IsNumeric(strColor) Then
        PlayerText(n) = CLng(strColor)
    End If
Next n

'load the nation list
If Not (rsNations.BOF And rsNations.EOF) Then
    rsNations.MoveFirst
    Nations.Clear
    While Not rsNations.EOF
        'drk@unxsoft.com : see the note in frmEmpCmd as to why I truncated the next line.
        Nations.AddNation rsNations.Fields("Number"), rsNations.Fields("Name") ', rsNations.Fields("yours"), rsNations.Fields("theirs")
        If rsNations.Fields("Name") = frmEmpCmd.CountryName Then
            CountryNumber = rsNations.Fields("Number")
        End If
        rsNations.MoveNext
    Wend
End If

Timer1.Enabled = False

'update lights on draw map
UpdateLights

'Put up the splash screen
On Error GoTo Timer1_Exit
frmDrawMap.Show
frmDrawMap.UpdateTimer.Interval = 1000

If Not frmDrawMap.MapValid Then
    frmDrawMap.picMap.AutoRedraw = True
    Set frmDrawMap.picMap.Picture = LoadPicture(App.path + "\WinACE.gif")
    frmDrawMap.picMap.Move (frmDrawMap.vsMap.Left - frmDrawMap.picMap(1).Width) / 2, _
                    (frmDrawMap.hsMap.top - frmDrawMap.picMap(1).Height) / 2
    frmDrawMap.picMap.Visible = True
End If


'If the report box is already up, need to bring it to the for-front
Dim frm As Form
For Each frm In Forms
    If frm.Name = "frmReport" Then
        On Error Resume Next  'Check for error in case box is just being created
        frmReport.SetFocus
        On Error GoTo 0
    End If
Next frm

Timer1_Exit:
Unload Me

End Sub


