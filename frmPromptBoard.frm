VERSION 5.00
Begin VB.Form frmPromptBoard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1905
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5760
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
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attacker"
      Height          =   855
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   1575
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Victim Ship"
      Height          =   855
      Left            =   1320
      TabIndex        =   6
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
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
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
      Left            =   4440
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Board"
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
      TabIndex        =   5
      Top             =   600
      Width           =   855
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
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form
'092803 rjk: Switched to SetUnitType and DisplayFirstSectorPanel
'            General reformatting.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'160705 rjk: Added ability to get OContent for Sea Sectors
'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

strCmd = "board " + txtTarget.Text + " "
If txtUnit.Visible Then
    strCmd = strCmd + txtUnit.Text
Else
    strCmd = strCmd + txtOrigin.Text
End If

frmEmpCmd.SubmitEmpireCommand strCmd, True

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetShips
If Not txtUnit.Visible Then
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
    GetOContent
    GetLandUnits
End If
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtUnit_GotFocus()
frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, True, True, False
End Sub

Private Sub txtTarget_GotFocus()
frmDrawMap.SetUnitDisplay udENEMY, TYPE_ALL, False, False, False
End Sub

Private Sub txtOrigin_GotFocus()
frmDrawMap.DisplayFirstSectorPanel
End Sub

