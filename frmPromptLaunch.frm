VERSION 5.00
Begin VB.Form frmPromptLaunch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1260
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5385
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
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtpath 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fire"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abort"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "target"
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
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Launch"
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
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092803 rjk: Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
'            General reformatting.
'180206 rjk: Replace pdump and lost with GetPlanes and GetLost.

Public strCmd As String
Public ShowRange As Boolean

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
strCmd = strCmd + " " + txtUnit.Text
frmEmpCmd.SubmitEmpireCommand "la1 " + txtPath.Text, False
frmEmpCmd.SubmitEmpireCommand strCmd, True

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetPlanes
GetLost
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
'rjk 8/11/03 Added LastLoad show the grid gets updated with inrange missiles each time the Launch form is launched
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtUnit_Change()
If Len(txtUnit.Text) = 0 Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub

Private Sub txtUnit_GotFocus()
If Len(txtPath.Text) > 0 And ShowRange Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_P_MISSILE, True, True, True, txtPath.Text
Else
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PC_LAUNCH, True, True, False
End If
End Sub
