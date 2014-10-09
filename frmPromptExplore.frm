VERSION 5.00
Begin VB.Form frmPromptExplore 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7440
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
      Left            =   7080
      TabIndex        =   15
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   3015
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmPromptExplore.frx":0000
         Left            =   1320
         List            =   "frmPromptExplore.frx":0016
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "radius"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "take path"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Military"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Civilian"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPromptExplore.frx":002C
      Left            =   3720
      List            =   "frmPromptExplore.frx":0042
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame frameMode 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3600
      TabIndex        =   16
      Top             =   720
      Width           =   1695
      Begin VB.OptionButton Mode 
         Caption         =   "Partial - Border"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Partial - Wheel"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Mode 
         Caption         =   "Complete"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Caption         =   "to each sector"
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
      Left            =   5520
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Width           =   375
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
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Explore"
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
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: removed dead variables
'090103 rjk: Added for additional exploration modes
'092803 rjk: General reformatting.
'101503 rjk: Disable the mode when in path mode
'101703 rjk: Added Strength fields to Sector database
'110703 rjk: Fixed a problem filling the path from the frmPromptPath form.
'112003 rjk: Added option to control strength updates
'122903 rjk: Fixed help to use the Title label
'270104 rjk: When selecting population, do not change the path option to radius
'050204 rjk: Select the radius option when the radius combo is selected.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'010804 rjk: Select the TakePath option when returning from the Path Selection form.
'250804 rjk: Disabled Civilians for 2K4.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim strUnit As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label1.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim strnum As String
Dim exploreMode As enumExploreMode '090103 rjk: Added mode for user to select exploration pattern

strnum = CStr(Val(Combo1.Text))
If Option3.Value Then

    'Compute text Destination
    strCmd = "explore " + strUnit + " " + txtOrigin.Text + " " + strnum + " " + txtPath.Text
    frmEmpCmd.SubmitEmpireCommand strCmd, True
    'database update
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
    frmEmpCmd.SubmitEmpireCommand "db2", False
Else
    If Mode(1).Value = True Then
        exploreMode = EXP_COMPLETE
    ElseIf Mode(2).Value = True Then
        exploreMode = EXP_BORDER
    ElseIf Mode(3).Value = True Then
        exploreMode = EXP_WHEEL
    Else 'should not happen
        exploreMode = EXP_COMPLETE
    End If
    ExploreOut txtOrigin.Text, strUnit, CInt(Combo2.Text), 1, 0, Val(Combo1.Text), exploreMode
End If

Unload Me
End Sub

Private Sub Combo1_Change()
txtOrigin.SetFocus
End Sub

Private Sub Combo1_Click()
txtOrigin.SetFocus
End Sub

Private Sub Combo2_Click()
Option4.Value = True
End Sub

Private Sub Combo2_Change()
Option4.Value = True
End Sub

Private Sub Form_Activate()
txtOrigin.SetFocus

Dim X As Integer
Dim sx As Integer
Dim sy As Integer

Combo2.Clear
For X = 1 To MaxExploreRadius
    Combo2.AddItem CStr(X)
Next X

Combo1.ListIndex = 0
Combo2.ListIndex = Combo2.ListCount - 1

'if there are no civs in the sector, default to mil
If ParseSectors(sx, sy, txtOrigin.Text) Then
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        If rsSectors.Fields("civ") < 2 Then
            Option2.Value = True
        End If
    End If
End If
If frmOptions.b2K4 Then
    Option1.Enabled = False
End If
If Len(txtPath.Text) > 0 Then
    Option3.Value = True
End If
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long   removed efj 8/2003
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 
Option1.Value = True
strUnit = "c"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
strUnit = "c"
txtOrigin.SetFocus
End Sub

Private Sub Option2_Click()
strUnit = "m"
txtOrigin.SetFocus
End Sub

Private Sub Option3_Click()
On Error Resume Next '101503 rjk: Removed not needed?
'110703 rjk: Added back. Used when filling txtPath the form the frmPromptPath form,
'            it is not the active form and can not accept the focus.
txtPath.SetFocus
'101503 rjk: Disable the Explore pattern frame, not appropriate for this mode
frameMode.Visible = False
End Sub

Private Sub Option4_Click()
'101503 rjk: Ensure the Explore pattern frame is displayed
frameMode.Visible = True
End Sub

Private Sub txtPath_Click()
Option3.Value = True
End Sub

Private Sub txtPath_Change()
Option3.Value = True
End Sub

Private Sub txtPath_DblClick()
'Make sure starting sector is valid before prompting
Dim sx As Integer
Dim sy As Integer
If Not ParseSectors(sx, sy, txtOrigin.Text) Then
    Exit Sub
End If

Load frmPromptPath
frmPromptPath.lblSector.Caption = txtOrigin.Text
frmPromptPath.lblEndSector.Caption = txtOrigin.Text
frmPromptPath.Caption = "Exploration Route"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless
End Sub
