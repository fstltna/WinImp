VERSION 5.00
Begin VB.Form frmPromptLook 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1905
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLimitMobility 
      Caption         =   "Limit the Mob. Gain to Max. Mob."
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbUpdates 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPromptLook.frx":0000
      Left            =   1680
      List            =   "frmPromptLook.frx":000D
      TabIndex        =   7
      Text            =   "1"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkIncludeUpdateEff 
      Caption         =   "Include Efficiency Gain from Update"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox chkIncludeUpdateMob 
      Caption         =   "Include Mobility Gain from Update"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
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
      Left            =   4080
      TabIndex        =   4
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblUpdates 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Updates"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Look"
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
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form
'092803 rjk: Made Okay the default button, added timestamp to ndump.
'            For marker only, only do the first ship and reset the unit field
'            Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
'            Switch OK to Show for marking, as OK does not leave form
'            General reformatting.
'100503 rjk: Removed timestamp ndump
'101403 rjk: Adding timestamp ndump
'            Added bmap to get X's on the map when doing lmine or mine
'102503 rjk: Added database update to unsail to get the UnitGrid to update
'270304 rjk: Switched over to DeleteAllRecords for clearing a table
'290404 rjk: Added the ability to include Update efficicency in NavMarkers
'100504 rjk: Added the ability to include Update mobility in NavMarkers
'060604 rjk: Changed the harbour check for NavMarkers to use the Bmap instead.
'120405 rjk: Added Update combo box for NavMarkers
'200405 rjk: Added checkbox for NavMarkers to control whether to limit
'            mobility to max. mobility or not.
'180206 rjk: Replace ndump with GetNukes.
'010108 rjk: Switch to new GetOrders function.

Public strCmd As String

Private Sub chkIncludeUpdateEff_Click()
If chkIncludeUpdateEff.Value <> vbUnchecked Or _
    chkIncludeUpdateMob.Value <> vbUnchecked Then
    cmbUpdates.Enabled = True
Else
    cmbUpdates.Enabled = False
End If
If chkIncludeUpdateMob.Value <> vbUnchecked Then
    chkLimitMobility.Enabled = True
Else
    chkLimitMobility.Enabled = False
End If
DrawMarkers
End Sub

Private Sub chkIncludeUpdateMob_Click()
If chkIncludeUpdateEff.Value <> vbUnchecked Or _
    chkIncludeUpdateMob.Value <> vbUnchecked Then
    cmbUpdates.Enabled = True
Else
    cmbUpdates.Enabled = False
End If
If chkIncludeUpdateMob.Value <> vbUnchecked Then
    chkLimitMobility.Enabled = True
Else
    chkLimitMobility.Enabled = False
End If
DrawMarkers
End Sub

Private Sub chkLimitMobility_Click()
DrawMarkers
End Sub

Private Sub cmbUpdates_Click()
DrawMarkers
End Sub

Private Sub cmdCancel_Click()
If strCmd = "marker" Then
    frmDrawMap.NavMarkerShip = vbNullString
    frmDrawMap.DrawHexes
End If

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
If strCmd = "marker" Then
    DrawMarkers
Else
    frmEmpCmd.SubmitEmpireCommand strCmd + " " + txtUnit.Text, True
    If strCmd = "look" Or strCmd = "llook" Or strCmd = "sonar" Then
        'just force the map to redraw
        frmEmpCmd.SubmitEmpireCommand "db1", False
        frmEmpCmd.SubmitEmpireCommand "db2", False
    ElseIf strCmd = "disarm" Then
        frmEmpCmd.SubmitEmpireCommand "db1", False
        GetNukes
        frmEmpCmd.SubmitEmpireCommand "db2", False
    ElseIf strCmd = "mine" Or strCmd = "lmine" Then
        frmEmpCmd.SubmitEmpireCommand "db1", False
        frmEmpCmd.SubmitEmpireCommand "bmap *", False '101403 rjk: Added bmap to get X's on the map
        frmEmpCmd.SubmitEmpireCommand "db2", False
    ElseIf strCmd = "unsail" Then '102503 rjk: Added unsail to get the UnitGrid to update
        GetOrders True
    End If
    
    frmDrawMap.DisplayFirstSectorPanel
    Unload Me
End If
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
If strCmd = "marker" Then '092803 rjk: Switch OK to Show for marking, as OK does not leave form
    cmdOK.Caption = "Show"
    lblUpdates.Visible = True
    cmbUpdates.Visible = True
    cmbUpdates.ListIndex = 0
    If frmDrawMap.bNavMarkerShipIncludeUpdateMob Or _
        frmDrawMap.bNavMarkerShipIncludeUpdateEff Then
        cmbUpdates.Enabled = True
        chkLimitMobility.Enabled = True
    Else
        cmbUpdates.Enabled = False
        chkLimitMobility.Enabled = False
    End If
    
    chkIncludeUpdateMob.Visible = True
    If frmDrawMap.bNavMarkerShipIncludeUpdateMob Then
        chkIncludeUpdateMob = vbChecked
        cmbUpdates.Enabled = True
    Else
        chkIncludeUpdateMob = vbUnchecked
    End If
    chkIncludeUpdateEff.Visible = True
    If frmDrawMap.bNavMarkerShipIncludeUpdateEff Then
        chkIncludeUpdateEff = vbChecked
    Else
        chkIncludeUpdateEff = vbUnchecked
    End If
    chkLimitMobility.Visible = True
    If frmDrawMap.bNavMarkerShipLimitMobility Then
        chkLimitMobility = vbChecked
    Else
        chkLimitMobility = vbUnchecked
    End If
Else
    cmdOK.Caption = "OK"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtUnit_Change()
If strCmd = "marker" Then
    DrawMarkers
End If
End Sub

Private Sub txtUnit_GotFocus()
Select Case strCmd
Case "look", "marker", "radar", "unsail"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "mine"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_S_MINE, True, True, False
Case "sonar"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_S_SONAR, True, True, False
Case "llook"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, True, False
Case "lradar"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_L_RADAR, False, True, False
Case "lmine"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_L_ENGINEER, True, True, False
Case "supply"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_L_SUPPLY, True, True, False
Case "disarm"
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_P_MISSILE, False, False, False
End Select
End Sub

Private Sub DrawMarkers()
Dim bEff As Boolean
Dim hldIndex As String

If InStrRev(txtUnit.Text, "/") > 0 Then '290404 rjk: only mark the last unit
    txtUnit.Text = Right$(txtUnit.Text, Len(txtUnit.Text) - InStrRev(txtUnit.Text, "/")) 'reset to only the last unit
End If
frmDrawMap.NavMarkerShip = txtUnit.Text
If chkIncludeUpdateMob <> vbUnchecked Then
    frmDrawMap.bNavMarkerShipIncludeUpdateMob = True
Else
    frmDrawMap.bNavMarkerShipIncludeUpdateMob = False
End If
    
bEff = False
hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"
If Len(txtUnit.Text) > 0 Then
    rsShip.Seek "=", Val(txtUnit.Text)
    If Not rsShip.NoMatch Then
        rsBmap.Seek "=", rsShip.Fields("x"), rsShip.Fields("y")
        If Not rsBmap.NoMatch Then
            If rsBmap.Fields("des") = "h" Then
                bEff = True
            End If
        End If
    End If
End If
rsShip.Index = hldIndex
If bEff Then
    chkIncludeUpdateEff.Enabled = True
Else
    chkIncludeUpdateEff.Value = vbUnchecked
    chkIncludeUpdateEff.Enabled = False
End If
If chkIncludeUpdateEff <> vbUnchecked Then
    frmDrawMap.bNavMarkerShipIncludeUpdateEff = True
Else
    frmDrawMap.bNavMarkerShipIncludeUpdateEff = False
End If

frmDrawMap.iNavMarkerShipUpdates = cmbUpdates.ItemData(cmbUpdates.ListIndex)
If chkLimitMobility <> vbUnchecked Then
    frmDrawMap.bNavMarkerShipLimitMobility = True
Else
    frmDrawMap.bNavMarkerShipLimitMobility = False
End If
frmDrawMap.DrawHexes
End Sub

