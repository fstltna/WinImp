VERSION 5.00
Begin VB.Form frmPromptFleet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1395
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5880
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
      TabIndex        =   7
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cmbSub 
      Height          =   315
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "to"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Add"
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
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Fleet"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPromptFleet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'120803 efj: removed dead variables
'280903 rjk: Added check to ensure the Fleet group is selected before
'            submitting command. Added timestamp to sdump.
'            Expand the form to work for planes and land units as well.
'            Added initial field selection. General reformatting.
'            Added Unit grid selection (SetUnitType, DisplayFirstSectorPanel)
'271003 rjk: Added cmbSub_Click to pick when the combo box selected from a click, needed to
'            enabled the Okay button if the user picks an existing fleet.
'110106 rjk: Fixed FleetAdd to not cancel the OK button whenever the map is touched.
'180206 rjk: Replace sdump, ldump, pdump with GetShips, GetLandUnits and GetPlanes.

Private Sub cmbSub_Change()
If cmbSub.Text = "" Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub

Private Sub cmbSub_Click()
If cmbSub.Text = "" Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

strCmd = LCase$(Label2.Caption) + " " + cmbSub.Text + " " + txtUnit.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If Label2 = "Fleet" Then '092503 rjk: Added dumps for appropriate fleet
    GetShips
ElseIf Label2 = "Wing" Then
    GetPlanes
Else
    GetLandUnits
End If

frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Form_Activate()
If Len(cmbSub.Text) > 0 And Len(txtUnit.Text) > 0 Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
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

Public Sub LoadCombo() ' 8/12/03 efj  removed, 09/11/03 rjk added back, used for FleetAdd from menu
Dim rs As Recordset
Dim strSubFleet As String
Dim Fleet As String

'Clear old records
strSubFleet = vbNullString
cmbSub.Clear

'Check recordset
If Label2 = "Fleet" Then
    Set rs = rsShip
ElseIf Label2 = "Wing" Then
    Set rs = rsPlane
Else
    Set rs = rsLand
End If

'Run thru list
rs.MoveFirst
While Not rs.EOF
    'If fleet has not been added
    Fleet = rs.Fields(4)
    If InStr(strSubFleet, Fleet) = 0 Then
        strSubFleet = strSubFleet + Fleet
        cmbSub.AddItem Fleet
    End If
rs.MoveNext
Wend
End Sub

Private Sub txtUnit_GotFocus()
If Label2 = "Fleet" Then
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
ElseIf Label2 = "Wing" Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
Else 'Army
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
End If
End Sub
