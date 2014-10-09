VERSION 5.00
Begin VB.Form frmPromptCensus 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1305
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmPromptCensus.frx":0000
      Left            =   240
      List            =   "frmPromptCensus.frx":0013
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
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
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Census"
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
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "for"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptCensus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Added Option Explicit and removed dead variables
'092703 rjk: General reformatting.  Added initial field selection.
'            Repositioned fields on form so label2 would not override label1.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'250606 rjk: Added start/stop unit support for 4.3.6 servers.
'181006 rjk: Ensure the unit list only appears for start and stop.

Private Sub FillForm()
Dim strCmd As String

strCmd = LCase$(Label2.Caption)

On Error Resume Next
If VersionCheck(4, 3, 6) = VER_LESS Or (strCmd <> "start" And strCmd <> "stop") Then
    cmbType.Visible = False
    cmbType.Text = ""
    txtOrigin.Visible = True
    txtUnit.Visible = False
    txtOrigin.SetFocus
    Exit Sub
End If
cmbType.Visible = True
If cmbType.ListIndex = 0 Then
    txtOrigin.Visible = True
    txtUnit.Visible = False
    txtOrigin.SetFocus
Else
    txtOrigin.Visible = False
    txtUnit.Visible = True
    txtUnit.SetFocus
End If
Select Case cmbType.ListIndex
Case 0
    frmDrawMap.DisplayFirstSectorPanel
Case 1
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case 2
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
Case 3
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
Case 4
    frmDrawMap.SetUnitDisplay udNUKE, TYPE_ALL, False, False, False
End Select
End Sub
Private Sub cmbType_Change()
FillForm
End Sub

Private Sub cmbType_Click()
FillForm
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

strCmd = LCase$(Label2.Caption)
'database update
If strCmd = "start" Or strCmd = "stop" Then
    If cmbType.ListIndex <= 0 Then
        frmEmpCmd.SubmitEmpireCommand strCmd + " " + LCase$(Left$(cmbType.Text, 4)) + _
            " " + txtOrigin.Text, True
    Else
        frmEmpCmd.SubmitEmpireCommand strCmd + " " + LCase$(Left$(cmbType.Text, 4)) + _
            " " + txtUnit.Text, True
    End If
    frmEmpCmd.SubmitEmpireCommand "db1", False
    Select Case (cmbType.ListIndex)
    Case -1, 0:
        GetSectors
        '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
        GetCurrentStrength tsSectors
    Case 1:
        GetShips
    Case 2:
        GetLandUnits
    Case 3:
        GetPlanes
    Case 4:
        GetNukes
    End Select
    frmEmpCmd.SubmitEmpireCommand "db2", False
ElseIf strCmd = "Import Offset" Then
Else
    frmEmpCmd.SubmitEmpireCommand strCmd + " " + txtOrigin.Text, True
End If

Unload Me
End Sub

Private Sub Form_Activate()
If (Label2.Caption = "Start" Or Label2.Caption = "Stop") And VersionCheck(4, 3, 6) <> VER_LESS Then
    FillForm
Else
    cmbType.Visible = False
    cmbType.Text = ""
    txtOrigin.Visible = True
    txtUnit.Visible = False
    txtOrigin.SetFocus '092203 rjk: added to ensure the sector map will update the txtOrigin field
End If
End Sub

Private Sub Form_Load()
'Make Stay always on top
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub
