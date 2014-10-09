VERSION 5.00
Begin VB.Form frmPromptFire 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7080
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
      Left            =   6720
      TabIndex        =   19
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkAbort 
      Caption         =   "Abort Remaining Shots on Return Fire"
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPromptFire.frx":0000
      Left            =   5160
      List            =   "frmPromptFire.frx":0013
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat Firing"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attacker"
      Height          =   1095
      Left            =   3240
      TabIndex        =   14
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target"
      Height          =   1095
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   1095
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option3 
         Caption         =   "fortress"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ship"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "land"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fire !"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Range to Target:"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Firing Range:"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
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
      Left            =   6000
      TabIndex        =   15
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Fire"
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
      TabIndex        =   12
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
      Left            =   960
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form
'092803 rjk: General reformatting. Switched to SetUnitType and
'            DisplayFirstSectorPanel
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace sdump and ldump with GetShips and GetLandUnits.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private Sub Check1_Click()
chkAbort.Enabled = (Check1.Value = vbChecked)
End Sub

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

If Option1 Then
    strCmd = " ship " + txtUnit.Text
ElseIf Option2 Then
    strCmd = " land " + txtUnit.Text
Else
    strCmd = " sect " + txtOrigin.Text
End If

strCmd = "fire " + strCmd + " " + txtTarget.Text

'Execute fire command
If Check1.Value = vbChecked Then
    num = Val(Combo1.Text)
Else
    num = 1
End If

If num > 1 Then
    'Check abort flag
    If chkAbort.Value = vbChecked Then
        frmEmpCmd.SubmitEmpireCommand "bfa", False
    Else
        frmEmpCmd.SubmitEmpireCommand "bf1", False
    End If
    For X = 1 To num
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    Next
    frmEmpCmd.SubmitEmpireCommand "bf2", False
Else
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If Option1.Value Then
    GetShips
ElseIf Option2.Value Then
    GetLandUnits
Else
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
End If
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
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

'Move unit box over origin box
txtUnit.Move txtOrigin.Left, txtOrigin.top, txtOrigin.Width, txtOrigin.Height
 
chkAbort.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
txtUnit.Visible = True
txtOrigin.Visible = False
SetFiringRange
txtUnit.SetFocus
End Sub

Private Sub Option2_Click()
txtUnit.Visible = True
txtOrigin.Visible = False
SetFiringRange
txtUnit.SetFocus
End Sub

Private Sub Option3_Click()
txtUnit.Visible = False
txtOrigin.Visible = True
SetFiringRange
txtOrigin.SetFocus
End Sub

Private Sub txtOrigin_Change()
SetFiringRange
SetTargetRange
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

Private Sub txtOrigin_GotFocus()
frmDrawMap.DisplayFirstSectorPanel
End Sub
Private Sub txtTarget_Change()
SetTargetRange
End Sub

Private Sub txtUnit_Change()
SetFiringRange
SetTargetRange
End Sub

Private Sub txtUnit_GotFocus()
If Option1 Then
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_SC_FIRE, False, True, False
End If
If Option2 Then
    frmDrawMap.SetUnitDisplay udLAND, TYPE_LC_FIRE, False, True, False
End If
End Sub

Private Sub SetFiringRange()
Dim frange As Single
Dim sx As Integer
Dim sy As Integer
Dim hldIndex As Variant

On Error GoTo SetFiringRangeError

If Option1 Then
    'Check for ship
    hldIndex = rsShip.Index
    rsShip.Index = "PrimaryKey"
    rsShip.Seek "=", CInt(txtUnit.Text)
    If Not rsShip.NoMatch Then
        frange = UnitFireRange(rsShip.Fields("tech"), rsShip.Fields("rng"))
    End If
    rsShip.Index = hldIndex
ElseIf Option2 Then
    'Check for land
    hldIndex = rsLand.Index
    rsLand.Index = "PrimaryKey"
    rsLand.Seek "=", CInt(txtUnit.Text)
    If Not rsLand.NoMatch Then
        frange = UnitFireRange(rsLand.Fields("tech"), rsLand.Fields("frg"))
    End If
    rsLand.Index = hldIndex
Else
    If ParseSectors(sx, sy, txtOrigin.Text) Then
        'Get Sector Information
        rsSectors.Seek "=", sx, sy
        If Not rsSectors.NoMatch Then
            frange = FortRange(TechLevel, rsSectors.Fields("eff").Value)
        Else
            frange = FortRange(TechLevel)
        End If
    End If
End If

SetFiringRangeError:
Label4.Caption = "Firing Range: " + Format(frange, "#0.00")
End Sub

Private Sub SetTargetRange()
' Dim frange As Single    8/12/03 efj  removed
Dim sx1 As Integer
Dim sy1 As Integer
Dim sx2 As Integer
Dim sy2 As Integer
Dim SetLabel As Boolean
Dim hldIndex As Variant

On Error GoTo SetTargetRangeError

If Len(txtTarget.Text) = 0 Then
    GoTo SetTargetRangeError
End If

SetLabel = True
If Option1 Then
    'Check for ship
    hldIndex = rsShip.Index
    rsShip.Index = "PrimaryKey"
    rsShip.Seek "=", CInt(txtUnit.Text)
    If Not rsShip.NoMatch Then
        sx1 = rsShip.Fields("x")
        sy1 = rsShip.Fields("y")
    Else
        SetLabel = False
    End If
    rsShip.Index = hldIndex
ElseIf Option2 Then
    'Check for land
    hldIndex = rsLand.Index
    rsLand.Index = "PrimaryKey"
    rsLand.Seek "=", CInt(txtUnit.Text)
    If Not rsLand.NoMatch Then
        sx1 = rsLand.Fields("x")
        sy1 = rsLand.Fields("y")
    Else
        SetLabel = False
    End If
    rsLand.Index = hldIndex
Else
    If Not ParseSectors(sx1, sy1, txtOrigin.Text) Then
        SetLabel = False
    End If
End If

'get the target sector
If Not ParseSectors(sx2, sy2, txtTarget.Text) Then
    'Check for known enemy unit
    hldIndex = rsEnemyUnit.Index
    rsEnemyUnit.Index = "ByID"
    rsEnemyUnit.Seek "=", "s", CInt(txtTarget.Text)
    If rsEnemyUnit.NoMatch Then
        SetLabel = False
    Else
        sx2 = rsEnemyUnit.Fields("x")
        sy2 = rsEnemyUnit.Fields("y")
    End If
    rsEnemyUnit.Index = hldIndex
End If

If SetLabel Then
    Label5.Caption = "Range to Target: " + CStr(SectorDistance(sx1, sy1, sx2, sy2))
Else
    Label5.Caption = "Range to Target: "
End If
Exit Sub

SetTargetRangeError:
Label5.Caption = "Range to Target: "
End Sub


