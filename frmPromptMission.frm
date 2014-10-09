VERSION 5.00
Begin VB.Form frmPromptMission 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtRange 
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      Height          =   1455
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "land"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ship"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "plane"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Set Reaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Op Sector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Mission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Mission"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptMission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'120803 efj: removed dead variables
'270903 rjk: general reformatting
'030206 rjk: request an update of the mission values when changing a mission.
'180206 rjk: Replace mission with xdump for 4.3.0 servers.

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim strType As String

If Option1.Value Then
    strType = "ship "
ElseIf Option2.Value Then
    If Len(Trim$(txtRange.Text)) > 0 Then
        strCmd = "lrange " + txtUnit.Text + " " + txtRange.Text
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
    strType = "land "
Else
    strType = "plane "
End If

strCmd = "mission " + strType + txtUnit.Text + " " + Left$(Combo1.Text, 1) + " " + txtOrigin + " " + txtNum
frmEmpCmd.SubmitEmpireCommand strCmd, True
If VersionCheck(4, 3, 0) = VER_LESS Then
    strCmd = "mission " + strType + txtUnit.Text + " q"
Else
    strCmd = "xdump " + strType + txtUnit.Text
End If
frmEmpCmd.SubmitEmpireCommand "db1", False
frmEmpCmd.SubmitEmpireCommand strCmd, False
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
Combo1.Clear
Combo1.AddItem "interdiction"
Combo1.AddItem "clear"
Combo1.AddItem "query"
Combo1.ListIndex = 0
Label6.Visible = False
txtRange.Visible = False
End Sub

Private Sub Option2_Click()
Combo1.Clear
Combo1.AddItem "interdiction"
Combo1.AddItem "reserve"
Combo1.AddItem "clear"
Combo1.AddItem "query"
Combo1.ListIndex = 0
Label6.Visible = True
txtRange.Visible = True
End Sub

Private Sub Option3_Click()
Combo1.Clear
Combo1.AddItem "interdiction"
Combo1.AddItem "support"
Combo1.AddItem "off support"
Combo1.AddItem "def support"
Combo1.AddItem "escort"
Combo1.AddItem "air defense"
Combo1.AddItem "clear"
Combo1.AddItem "query"
Combo1.ListIndex = 0
Label6.Visible = False
txtRange.Visible = False
End Sub
