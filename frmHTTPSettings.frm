VERSION 5.00
Begin VB.Form frmHTTPSettings 
   Caption         =   "HTTP Proxy Settings"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame fraProxySettings 
      Caption         =   "Proxy Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3855
      Begin VB.CheckBox chkProxySettings 
         Caption         =   "Use Proxy Settings"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtProxyPassword 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Leave Blank if not needed"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtProxyServerName 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Proxy Server Name"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtProxyUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Leave Blank if not needed"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblProxyPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label lblProxyServerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Server"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblProxyUser 
         Alignment       =   1  'Right Justify
         Caption         =   "User"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame fraXMLRPC 
      Caption         =   "XML RPC"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtXMLRPCServerParameters 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtXMLRPCServerName 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblXMLRPCServerParameters 
         Alignment       =   1  'Right Justify
         Caption         =   "Parameters"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblXMLRPCServerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Server"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmHTTPSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'080605 rjk: Added XML RPC Parameters - http proxy through firewalls

Public strXMLRPCServerName As String
Public strXMLRPCServerParameters As String
Public bProxyEnabled As Boolean
Public strProxyServerName As String
Public strProxyUser As String
Public strProxyPassword As String

Private Sub chkProxySettings_Click()
If chkProxySettings.Value <> vbUnchecked Then
    txtProxyServerName.Enabled = True
    txtProxyUser.Enabled = True
    txtProxyPassword.Enabled = True
Else
    txtProxyServerName.Enabled = False
    txtProxyUser.Enabled = False
    txtProxyPassword.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
strXMLRPCServerName = txtXMLRPCServerName
strXMLRPCServerParameters = txtXMLRPCServerParameters
If chkProxySettings.Value <> vbUnchecked Then
    bProxyEnabled = True
Else
    bProxyEnabled = False
End If
strProxyServerName = txtProxyServerName.Text
strProxyUser = txtProxyUser.Text
strProxyPassword = txtProxyPassword.Text
SaveXMLRPCParameters
Unload Me
End Sub

Private Sub Form_Load()
txtXMLRPCServerName = strXMLRPCServerName
txtXMLRPCServerParameters = strXMLRPCServerParameters
If bProxyEnabled Then
    chkProxySettings.Value = vbChecked
Else
    chkProxySettings.Value = vbUnchecked
End If
txtProxyServerName.Text = strProxyServerName
txtProxyUser.Text = strProxyUser
txtProxyPassword.Text = strProxyPassword
chkProxySettings_Click
End Sub

Public Sub LoadXMLRPCParameters()
bXMLRPC = GetSetting(APPNAME, "XML RPC Options", "Use XMLRPC ", False)
strXMLRPCServerName = GetSetting(APPNAME, "XML RPC Options", "Server Name", "http://emproxy.tryba.nl")
strXMLRPCServerParameters = GetSetting(APPNAME, "XML RPC Options", "Server Parameters", "empire.servlets/servlet/XmlRpc")
bProxyEnabled = GetSetting(APPNAME, "XML RPC Options", "Proxy Enabled", False)
strProxyServerName = GetSetting(APPNAME, "XML RPC Options", "Proxy Server Name", "")
strProxyUser = GetSetting(APPNAME, "XML RPC Options", "Proxy User", "")
strProxyPassword = GetSetting(APPNAME, "XML RPC Options", "Proxy Password", "")
End Sub

Public Sub SaveXMLRPCParameters()
SaveSetting APPNAME, "XML RPC Options", "Use XMLRPC ", bXMLRPC
SaveSetting APPNAME, "XML RPC Options", "Server Name", strXMLRPCServerName
SaveSetting APPNAME, "XML RPC Options", "Server Parameters", strXMLRPCServerParameters
SaveSetting APPNAME, "XML RPC Options", "Proxy Enabled", bProxyEnabled
SaveSetting APPNAME, "XML RPC Options", "Proxy Server Name", strProxyServerName
SaveSetting APPNAME, "XML RPC Options", "Proxy User", strProxyUser
SaveSetting APPNAME, "XML RPC Options", "Proxy Password", strProxyPassword
End Sub

