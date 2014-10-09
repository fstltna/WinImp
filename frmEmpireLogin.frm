VERSION 5.00
Begin VB.Form frmEmpireLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empire Login"
   ClientHeight    =   5760
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3403.2
   ScaleMode       =   0  'User
   ScaleWidth      =   4295.677
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Server Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   4335
      Begin VB.ComboBox cmbThemeGame 
         Height          =   315
         ItemData        =   "frmEmpireLogin.frx":0000
         Left            =   1200
         List            =   "frmEmpireLogin.frx":0002
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   900
         Width           =   3015
      End
      Begin VB.CommandButton cmdHTTP 
         Caption         =   "Settings"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkHTTP 
         Caption         =   "HTTP"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Allows the use of Unicode (Multiple Languages) on the game server"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkSanc 
         Caption         =   "List Sanctuaries"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkKill 
         Caption         =   "Kill Hung Session"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkUTF8 
         Caption         =   "UTF-8"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         ToolTipText     =   "Allows the use of Unicode (Multiple Languages) on the game server"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblThemeGame 
         Caption         =   "Theme Game"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOffline 
      Caption         =   "Offline"
      Height          =   390
      Left            =   1680
      TabIndex        =   21
      Top             =   5280
      Width           =   1260
   End
   Begin VB.TextBox txtUser 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   3165
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client Options"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   4335
      Begin VB.CheckBox chkDeity 
         Caption         =   "Sign on as Deity"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "AutoRead"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "AutoUpdate"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   960
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1260
   End
   Begin VB.TextBox txtPort 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtHostName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   3165
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3240
      TabIndex        =   11
      Top             =   5280
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   3165
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User:"
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblGame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   112.673
      X2              =   4394.266
      Y1              =   354.5
      Y2              =   354.5
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Port Number"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Host Name:"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Country Name:"
      Height          =   270
      Index           =   2
      Left            =   105
      TabIndex        =   0
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   3
      Left            =   105
      TabIndex        =   12
      Top             =   1920
      Width           =   1080
   End
End
Attribute VB_Name = "frmEmpireLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub chkHTTP_Click()
If chkHTTP.Value <> vbUnchecked Then
    cmdHTTP.Enabled = True
Else
    cmdHTTP.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

'set the global var to false
'to denote a failed login
If bXMLRPC Then
    xmlrpcServer.CloseConnection
    Set xmlrpcServer = Nothing
Else
    frmEmpCmd.Winsock.Close
End If
ConnectedtoHost = False
Me.Hide
End Sub

Private Sub cmdConnect_Click()
Dim filenumber As Integer
'Dim strDeity As Integer 101603 Switched to string based on usage
Dim strDeity As String
'101603 rjk: Added to store the current setting of AutoUpdate and AutoRead
Dim strAutoUpdate As String
Dim strAutoRead As String
Dim strUTF8 As String

'Check that all data is entered
If Len(txtHostName.Text) = 0 Or _
    Len(txtPort.Text) = 0 Or _
    (Len(txtUserName.Text) = 0 And chkSanc.Value = vbUnchecked) Or _
    (Len(txtUser.Text) = 0 And chkSanc.Value = vbUnchecked) Or _
    (Len(txtPassword.Text) = 0 And chkSanc.Value = vbUnchecked) Then
    Label1.Caption = "Error: Must fill each field"
    Exit Sub
End If

'Set Deity Flag
If chkDeity.Value = vbChecked Then
    strDeity = "1"
Else
    strDeity = "0"
End If

'101603 rjk: Creating string to store current setting of AutoUpdate and AutoRead
If chkUpdate.Value = vbChecked Then
    strAutoUpdate = "1"
Else
    strAutoUpdate = "0"
End If
If chkAutoRead.Value = vbChecked Then
    strAutoRead = "1"
Else
    strAutoRead = "0"
End If
If chkUTF8.Value <> vbUnchecked Then
    strUTF8 = "1"
Else
    strUTF8 = "0"
End If

If frmEmpireLogin.chkHTTP <> vbUnchecked Then
    bXMLRPC = True
Else
    bXMLRPC = False
End If
frmHTTPSettings.SaveXMLRPCParameters

SaveSetting APPNAME, "Options", "Star Wars", False
SaveSetting APPNAME, "Options", "Retro", False
SaveSetting APPNAME, "Options", "2K4", False
SaveSetting APPNAME, "Options", "EE9", False
SaveSetting APPNAME, "Options", "SP2005", False
SaveSetting APPNAME, "Options", "SPAtlantis", False
SaveSetting APPNAME, "Options", "HeavyMetal", False
Select Case cmbThemeGame.ListIndex
Case 1:
    SaveSetting APPNAME, "Options", "Star Wars", True
Case 2:
    SaveSetting APPNAME, "Options", "Retro", True
Case 3:
    SaveSetting APPNAME, "Options", "2K4", True
Case 4:
    SaveSetting APPNAME, "Options", "EE9", True
Case 5:
    SaveSetting APPNAME, "Options", "SP2005", True
Case 6:
    SaveSetting APPNAME, "Options", "SPAtlantis", True
Case 7:
    SaveSetting APPNAME, "Options", "HeavyMetal", True
End Select

'Save connection data
filenumber = FreeFile   ' Get unused file number.
If VerifySubDirectory("Games", True) Then
    Open App.path + "\Games\" + GameName + ".egp" For Output As #filenumber
    '101603 rjk: Added storing current setting of AutoUpdate and AutoRead
    Write #filenumber, Trim$(txtHostName.Text), Trim$(txtPort.Text), txtUserName.Text, txtPassword.Text, txtUser.Text, strDeity, strAutoUpdate, strAutoRead, strUTF8
    Close #filenumber
End If

On Error GoTo Connect_Error

'Make the connection to the Empire Server
If Not ConnectedtoHost Then
    If bXMLRPC Then
        On Error GoTo XMLRPC_Error
        Set xmlrpcServer = New EmpXMLRPC
        xmlrpcServer.OpenConnection Trim$(txtHostName.Text), Val(txtPort)
    Else
        frmEmpCmd.Winsock.RemoteHost = Trim$(txtHostName.Text)
        frmEmpCmd.Winsock.RemotePort = Val(txtPort)
        frmEmpCmd.Winsock.Connect
    End If
    
    'Set the timer to check for the Winsock State
    Timer1.Interval = 350
    Timer1.Enabled = True
    Timer1_Timer
Else
    Me.Hide 'Your done - must be reconnect.
End If
Exit Sub

Connect_Error:
Label1.Caption = "Error: " + Err.Description
Exit Sub

XMLRPC_Error:
Label1.Caption = "Error: " + Err.Description
MsgBox "You need to load Microsoft XML Parser 4.0 before you can HTTP - XML RPC capability", vbOKOnly
End Sub

Public Sub go()
Dim filenumber As Long
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s4 As String

frmHTTPSettings.LoadXMLRPCParameters
If bXMLRPC Then
    chkHTTP = vbChecked
    cmdHTTP.Enabled = True
Else
    chkHTTP = vbUnchecked
    cmdHTTP.Enabled = False
End If
'Check for game name file - if not found, then must be a new entry
If Len(Dir(App.path + "\Games\" + GameName + ".egp")) > 0 Then
    'Load the parameters from the current file
    filenumber = FreeFile   ' Get unused file number.
    Open App.path + "\Games\" + GameName + ".egp" For Input As #filenumber
    Input #filenumber, s1, s2, s3, s4
    txtHostName.Text = s1
    txtPort.Text = s2
    txtUserName.Text = s3
    txtPassword.Text = s4
    On Error Resume Next
    Input #filenumber, s1
    txtUser.Text = s1
    Input #filenumber, s1
    If s1 = "1" Then
        chkDeity.Value = vbChecked
    Else
        chkDeity.Value = vbUnchecked
    End If
    '101603 rjk: Retrieving of the AutoUpdate and AutoRead for each game
    If Not EOF(filenumber) Then
        Input #filenumber, s1, s2
        If s1 = "1" Then
            chkUpdate = vbChecked
        Else
            chkUpdate = vbUnchecked
        End If
        If s2 = "1" Then
            chkAutoRead = vbChecked
        Else
            chkAutoRead = vbUnchecked
        End If
    End If
    If Not EOF(filenumber) Then
        Input #filenumber, s1
        If s1 = "1" Then
            chkUTF8 = vbChecked
        Else
            chkUTF8 = vbUnchecked
        End If
    End If
    Close #filenumber
End If

cmbThemeGame.AddItem "None"
cmbThemeGame.AddItem "Star Wars"
cmbThemeGame.AddItem "Retro"
cmbThemeGame.AddItem "2K4"
cmbThemeGame.AddItem "EE9"
cmbThemeGame.AddItem "SP2005"
cmbThemeGame.AddItem "SPAtlantis"
cmbThemeGame.AddItem "HeavyMetal"

If GetSetting(APPNAME, "Options", "Star Wars", False) Then
    cmbThemeGame.ListIndex = 1
ElseIf GetSetting(APPNAME, "Options", "Retro", False) Then
    cmbThemeGame.ListIndex = 2
ElseIf GetSetting(APPNAME, "Options", "2K4", False) Then
    cmbThemeGame.ListIndex = 3
ElseIf GetSetting(APPNAME, "Options", "EE9", False) Then
    cmbThemeGame.ListIndex = 4
ElseIf GetSetting(APPNAME, "Options", "SP2005", False) Then
    cmbThemeGame.ListIndex = 5
ElseIf GetSetting(APPNAME, "Options", "SPAtlantis", False) Then
    cmbThemeGame.ListIndex = 6
ElseIf GetSetting(APPNAME, "Options", "HeavyMetal", False) Then
    cmbThemeGame.ListIndex = 7
Else
    cmbThemeGame.ListIndex = 0
End If

lblGame.Caption = "Game: " + GameName
End Sub

Private Sub cmdHTTP_Click()
frmHTTPSettings.Show vbModal
End Sub

Private Sub cmdOffline_Click()
bOffline = True
cmdConnect_Click
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Only unload if form Draw says OK
If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
    If Not frmDrawMap.ShutDown Then
        Cancel = 1
        Me.WindowState = vbMinimized
    End If
End If
End Sub

Private Sub Timer1_Timer()
Dim strState As String

If bXMLRPC Then
    If xmlrpcServer Is Nothing Then
        Label1.Caption = "Not Connected"
    ElseIf xmlrpcServer.Status Then
        Label1.Caption = "Connected"
    Else
        Label1.Caption = "Not Connected"
    End If
    
    'Check that existing connection is OK
    If ConnectedtoHost Then
        If xmlrpcServer.Status Then
            Exit Sub
        Else
            ConnectedtoHost = False
            frmEmpCmd.ConnectedtoHost = False
            frmEmpCmd.CheckLostConnection
        End If
    End If
    
    If xmlrpcServer Is Nothing Then
        If frmEmpCmd.bRelogin Then
            cmdConnect_Click
        End If
    ElseIf xmlrpcServer.Status Then
        If Not ConnectedtoHost Then
            SuccessfullConnection
        End If
    Else
        If frmEmpCmd.bRelogin Then
            cmdConnect_Click
        End If
    End If
    If xmlrpcServer Is Nothing Then
        Label1.Caption = "Not Connected"
    ElseIf xmlrpcServer.Status Then
        Label1.Caption = "Connected"
    Else
        Label1.Caption = "Not Connected"
    End If
Else
        'Check that existing connection is OK
        If ConnectedtoHost Then
            If frmEmpCmd.Winsock.State = sckConnected Then
                Exit Sub
            Else
                If frmEmpCmd.Winsock.State = sckClosing Then
                    frmEmpCmd.Winsock.Close
                End If
                ConnectedtoHost = False
                frmEmpCmd.ConnectedtoHost = False
                frmEmpCmd.CheckLostConnection
            End If
        End If
        
        'Check the state of the Winsock Connection
        Select Case frmEmpCmd.Winsock.State
            Case sckClosed:
                strState = "Closed"
                If frmEmpCmd.bRelogin Then
                    cmdConnect_Click
                End If
            Case sckOpen:
                strState = "Open"
            Case sckListening
                strState = "Listening"
            Case sckConnectionPending:
                strState = "Connection pending"
            Case sckResolvingHost:
                strState = "Resolving Host"
            Case sckHostResolved:
                strState = "Host resolved"
            Case sckConnecting:
                strState = "Connecting"
            Case sckConnected:
                strState = "Connected"
                If Not ConnectedtoHost Then
                    SuccessfullConnection
                End If
            Case sckClosing:
                strState = "Peer is closing the connection"
                frmEmpCmd.Winsock.Close
            Case sckError:
                strState = "Error"
                BadConnection
            Case Else
                strState = "Unknown - " + CStr(frmEmpCmd.Winsock.State)
        End Select
        
        Label1.Caption = strState
End If
End Sub

Private Sub SuccessfullConnection()
ConnectedtoHost = True
Timer1.Interval = 5000
Me.Hide
If frmEmpCmd.bRetry Or frmEmpCmd.bRelogin Then
    frmEmpCmd.bRelogin = False
    frmEmpCmd.LoginIntoServer
End If
End Sub

Private Sub BadConnection()
'092603 rjk: Added Cancel to work offline
'MsgBox "Error - Connection not Established", vbOKOnly, "Connection Error"
'Timer1.Enabled = False
'frmEmpCmd.Winsock.Close
Dim Reply As Integer

If bOffline Then
    Exit Sub
End If

Reply = MsgBox("Error - Connection not Established, use Cancel to work offline", vbOKCancel, "Connection Error")
If Reply <> vbCancel Then
    Timer1.Enabled = False
    If bXMLRPC Then
        xmlrpcServer.CloseConnection
        Set xmlrpcServer = Nothing
    Else
        frmEmpCmd.Winsock.Close
    End If
    bOffline = False
Else
    bOffline = True
    ConnectedtoHost = False
    frmEmpCmd.ConnectedtoHost = False
    Timer1.Enabled = False
    Me.Hide
End If
End Sub

