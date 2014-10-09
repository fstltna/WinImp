VERSION 5.00
Begin VB.Form frmPromptSail 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "cond"
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
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "path"
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sail"
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
End
Attribute VB_Name = "frmPromptSail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'091803 rjk: Added Unit grid selection when activating this form or selecting fields
'            and when disactivating return to standard Sector display.
'            General reformatting. Added setting the initial field on start up.
'092503 rjk: Added timestamp to ndump.
'100503 rjk: Removed timestamp ndump.
'101403 rjk: Adding timestamp ndump, add lost * to pick nukes being attached to missiles.
'101603 rjk: Update the map with info satellite report
'270304 rjk: Switched over to DeleteAllRecords for clearing a table
'180206 rjk: Replace ldump, pdump, sdump, ndump, and lost with GetLandUnits, GetPlanes,
'            GetShips, GetNukes and GetLost.
'010108 rjk: Switch to new GetOrders function.

Public strCmd As String

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
strCmd = strCmd + " " + txtUnit.Text + " " + txtPath.Text
If Combo1.Visible Then
    strCmd = strCmd + " " + Left$(Combo1.Text, 1)
End If
frmEmpCmd.SubmitEmpireCommand strCmd, True
'102503 rjk: Added database updates
frmEmpCmd.SubmitEmpireCommand "db1", False
Select Case Left$(strCmd, 3)
Case "arm"
    '101403 rjk: To pick up arm activity which remove nukes and attaches them to missiles.
    GetPlanes
    GetLost
    GetNukes
'Case "sat"
Case "sai"
    GetOrders True
Case "ret", "nam", "mqu"
    GetShips
Case "ran", "har"
    GetPlanes
Case "lra", "for", "mor", "lre"
    GetLandUnits
End Select
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

Private Sub txtPath_DblClick()
On Error GoTo txtPath_DblClick_Exit

'if this is a sail command - path is not valid
If (Not (Label2.Caption = "Sail" Or Label2.Caption = "Retreat")) Or Len(txtUnit.Text) = 0 Then
    Exit Sub
End If

'Make sure starting unit strind is valid before prompting
Dim units As Variant
Dim strOrigin As String
Dim hldIndex As Variant

strOrigin = vbNullString
units = ParseUnitString(txtUnit.Text)
If Not IsArray(units) Then
    Exit Sub
End If

'    CurrUnit = GetUnitList(txtUnit.Text, "p")
'    hldIndex = rsPlane.Index
'    rsPlane.Index = "PrimaryKey"
'    n = LBound(CurrUnit) + 1

If Left$(strCmd, 1) = "l" Then
    'find the first land in unit string - if not found, exit
    hldIndex = rsLand.Index
    rsLand.Index = "PrimaryKey"
    rsLand.Seek "=", units(LBound(units) + 1)
    If Not rsLand.NoMatch Then
        strOrigin = SectorString(rsLand.Fields("x"), rsLand.Fields("y"))
    End If
    rsLand.Index = hldIndex
    
Else
    'find the first ship in unit string - if not found, exit
    hldIndex = rsShip.Index
    rsShip.Index = "PrimaryKey"
    rsShip.Seek "=", units(LBound(units) + 1)
    If Not rsShip.NoMatch Then
        strOrigin = SectorString(rsShip.Fields("x"), rsShip.Fields("y"))
    End If
    rsShip.Index = hldIndex
    
End If

If Not Len(strOrigin) > 0 Then
    Exit Sub
End If

'setup path prompt
Load frmPromptPath
frmPromptPath.lblSector.Caption = strOrigin
frmPromptPath.lblEndSector.Caption = strOrigin
frmPromptPath.Caption = "Exploration Route"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless

txtPath_DblClick_Exit:
End Sub

Private Sub txtUnit_GotFocus()
Select Case strCmd
Case "sail", "mquota", "name", "retreat"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "fortify", "morale", "lrange", "lretreat"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
Case "range"
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
Case "satellite"
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_P_SATELLITE, True, True, False
Case "arm", "harden"
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_P_MISSILE, False, False, False
End Select
End Sub
