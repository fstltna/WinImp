VERSION 5.00
Begin VB.Form frmPromptTorp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1905
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7230
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
      Left            =   6840
      TabIndex        =   12
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPromptTorp.frx":0000
      Left            =   5640
      List            =   "frmPromptTorp.frx":0013
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat Shot"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attacker"
      Height          =   855
      Left            =   1920
      TabIndex        =   9
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target"
      Height          =   855
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fire !"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "at"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "times"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Torpedo"
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
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "from"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptTorp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'091803 rjk: Added Unit grid selection when activating this form or selecting fields
'            and when disactivating return to standard Sector display.
'            General reformatting. Added setting the initial field on start up.
'180206 rjk: Replace sdump with GetShips.

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim num As Integer
Dim X As Integer

strCmd = "torpedo " + txtUnit.Text + " " + txtTarget.Text

'Execute fire command
If Check1.Value = vbChecked Then
    num = Val(Combo1.Text)
Else
    num = 1
End If

If num > 1 Then
    frmEmpCmd.SubmitEmpireCommand "bf1", False
    For X = 1 To num
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    Next
    frmEmpCmd.SubmitEmpireCommand "bf2", False
Else
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetShips
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Combo1_Change()
Check1.Value = vbChecked
End Sub

Private Sub Combo1_Click()
Check1.Value = vbChecked
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
txtUnit.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtUnit_GotFocus()
frmDrawMap.SetUnitDisplay udSHIP, TYPE_S_TORP, True, True, False
End Sub

Private Sub txtTarget_GotFocus()
frmDrawMap.SetUnitDisplay udENEMY, TYPE_ALL, False, False, False
End Sub
