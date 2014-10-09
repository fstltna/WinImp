VERSION 5.00
Begin VB.Form frmPromptWarload 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1260
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5535
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
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Warload"
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
Attribute VB_Name = "frmPromptWarload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'091803 rjk: Added Unit grid selection when activating this form
'            and when disactivating Return to standard Sector display.
'            Also general reformatting
'101703 rjk: Added Strength fields to Sector database
'111403 rjk: Added so all the commands are combined on the command display
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'011506 rjk: Set the OK button to default.
'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Public strCmd As String
' Public strSector As String    8/2003 efj  removed

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
'111403 rjk: Added so all the commands are combined on the command display
frmEmpCmd.SubmitEmpireCommand "bf1", False

If strCmd = "swarload" Then
    frmEmpCmd.SubmitEmpireCommand "load mil " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "load gun " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "load shell " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "load food " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "load oil " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "load pet " + txtUnit.Text + " 999", True
End If

If strCmd = "lwarload" Then
    frmEmpCmd.SubmitEmpireCommand "lload mil " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "lload gun " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "lload shell " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "lload food " + txtUnit.Text + " 999", True
    frmEmpCmd.SubmitEmpireCommand "lload oil " + txtUnit.Text + " 999", True
End If
'111403 rjk: Added so all the commands are combined on the command display
frmEmpCmd.SubmitEmpireCommand "bf2", False

'refresh database
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors

If strCmd = "swarload" Then
    GetShips
End If

If strCmd = "lwarload" Then
    GetLandUnits
End If

frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtUnit_GotFocus()
Select Case strCmd
Case "swarload"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "lwarload"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
End Select
End Sub
