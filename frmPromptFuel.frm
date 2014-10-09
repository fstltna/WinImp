VERSION 5.00
Begin VB.Form frmPromptFuel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6735
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
      Left            =   6360
      TabIndex        =   11
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Oiler (optional)"
      Height          =   855
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
      Begin VB.TextBox txtUnit3 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ships"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "lands"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Fuel"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "with"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "fuel Units"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form
'092803 rjk: Switched to SetUnitType and DisplayFirstSectorPanel.
'            General reformatting.  Select Unit field when changing
'            to option2 or option3
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace ldump, sdump with GetLandUnits and GetShips.
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
If Option1.Value Then
    strCmd = strCmd + "fuel ship "
Else
    strCmd = strCmd + "fuel land "
End If

strCmd = strCmd + txtUnit.Text + " " + txtNum.Text + " " + txtUnit3.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'frmEmpCmd.SubmitEmpireCommand "dump " + strSector, False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
If Option1.Value Then
    GetShips
Else
    GetLandUnits
End If
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
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

Private Sub Option2_Click()
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
End Sub

Private Sub Option1_Click()
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
End Sub

Private Sub txtUnit_GotFocus()
If Option1.Value Then
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
ElseIf Option2.Value Then
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
End If
End Sub

Private Sub txtUnit3_GotFocus()
frmDrawMap.SetUnitDisplay udSHIP, TYPE_S_OILER, False, False, False
End Sub
