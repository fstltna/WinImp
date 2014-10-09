VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Range Table"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmShip 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox cmbShip 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Movement"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Range in Sectors"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Max Mob "
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblShipMax 
         Alignment       =   1  'Right Justify
         Caption         =   "00.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Per Turn "
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblShipTurn 
         Alignment       =   1  'Right Justify
         Caption         =   "00.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Sector Move Cost"
         Height          =   495
         Left            =   960
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblShipMoveCost 
         Alignment       =   1  'Right Justify
         Caption         =   "00.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblShipRange 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Firing Range"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame frmNuke 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox chkRoundTrip 
         Caption         =   "Round Trip"
         Height          =   375
         Left            =   840
         TabIndex        =   39
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optGroundBurst 
         Caption         =   "Ground Burst"
         Height          =   195
         Left            =   1800
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optAirBurst 
         Caption         =   "Air Burst"
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbNuke 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblNukeType 
         Alignment       =   1  'Right Justify
         Caption         =   "Nuke Type"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblTarget 
         Alignment       =   1  'Right Justify
         Caption         =   "Target"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame frmLand 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox cmbLand 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Firing Range"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLandRange 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame frmSectors 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   3495
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   3120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1320
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label14 
         Caption         =   "Coastwatch"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCoastwatch 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Radar Range"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRadar 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "100% Radar Station"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "60-100%"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblFort100 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "01-59%"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblFort59 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Fort Firing Range"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2520
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.VScrollBar vsTech 
      Height          =   495
      Left            =   1920
      Max             =   0
      Min             =   450
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtTech 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2520
      Width           =   615
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sectors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ships"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Land"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nukes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Airports"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTech 
      Alignment       =   2  'Center
      Caption         =   "Tech (Estimate)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmToolRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'100303 rjk: Protect against too tech level for the Vscroll bar.
'100304 rjk: Protect vsTech Min/Max against too high tech level.
'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
'            ships with SWEEP capabilities.
'160705 rjk: Fixed LoadComboBox to deal with empty build table.
'010106 rjk: Added Airport Range.
'210506 rjk: Fixed Nuke combobox, broken when airport was added.

Dim bscroll As Boolean
Dim ShipMobTurn As Integer
Dim ShipMobMax As Integer

Private Sub chkRoundTrip_Click()
CalculatePlaneRange vsTech.Value
End Sub

Private Sub cmbLand_Change()
CalculateLandRanges CSng(vsTech.Value)
End Sub

Private Sub cmbLand_Click()
CalculateLandRanges CSng(vsTech.Value)
End Sub

Private Sub cmbNuke_Change()
If TabStrip1.SelectedItem.Index = 4 Then
    CalculateNukeDamage
Else
    CalculatePlaneRange vsTech.Value
End If
End Sub

Private Sub cmbNuke_Click()
If TabStrip1.SelectedItem.Index = 4 Then
    CalculateNukeDamage
Else
    CalculatePlaneRange vsTech.Value
End If
End Sub

Private Sub cmbShip_Change()
CalculateShipRanges CSng(vsTech.Value)
End Sub

Private Sub cmbShip_Click()
CalculateShipRanges CSng(vsTech.Value)
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
bscroll = False

'100303 rjk: Protect against too tech level for the Vscroll bar.
If TechLevel > vsTech.Min Then 'scrollbar is setup in reverse direction.
    vsTech.Value = vsTech.Min
ElseIf TechLevel < vsTech.Max Then
    vsTech.Value = vsTech.Max
Else
    vsTech.Value = TechLevel
End If

TabStrip1_Click
CalculateSectorRanges TechLevel
LoadComboBox Me.cmbShip, "s"
LoadComboBox Me.cmbLand, "l"
LoadComboBox Me.cmbNuke, "n"

If Not (rsVersion.BOF And rsVersion.EOF) Then
    ShipMobMax = rsVersion.Fields("Max mob - Ship")
    ShipMobTurn = rsVersion.Fields("Mob gain - Ship")
End If

Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Len(NukeDamageType) > 0 Or Len(PlaneRangeType) > 0 Then
    NukeDamageType = vbNullString
    PlaneRangeType = vbNullString
    frmDrawMap.DrawHexes
End If
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub optAirBurst_Click()
CalculateNukeDamage
End Sub

Private Sub optGroundBurst_Click()
CalculateNukeDamage
End Sub

Private Sub TabStrip1_Click()
frmSectors.Visible = False
frmShip.Visible = False
frmLand.Visible = False
frmNuke.Visible = False
txtTech.Visible = True
lblTech.Visible = True
vsTech.Visible = True
Select Case (TabStrip1.SelectedItem.Index)
Case 1
    frmSectors.Visible = True
Case 2
    frmShip.Visible = True
Case 3
    frmLand.Visible = True
Case 4
    txtTech.Visible = False
    lblTech.Visible = False
    vsTech.Visible = False
    frmNuke.Visible = True
    optAirBurst.Visible = True
    optGroundBurst.Visible = True
    lblNukeType.Caption = "Nuke Type"
    chkRoundTrip.Visible = False
    LoadComboBox cmbNuke, "n"
Case 5
    txtTech.Visible = True
    lblTech.Visible = True
    vsTech.Visible = True
    frmNuke.Visible = True
    optAirBurst.Visible = False
    optGroundBurst.Visible = False
    lblNukeType.Caption = "Plane Type"
    chkRoundTrip.Visible = True
    LoadComboBox cmbNuke, "p"
End Select
CalculateRanges CSng(vsTech.Value)
End Sub

Private Sub txtOrigin_Change()
If TabStrip1.SelectedItem.Index = 4 Then
    CalculateNukeDamage
Else
    CalculatePlaneRange vsTech.Value
End If
End Sub

Private Sub txtTech_Change()
If Not bscroll Then
    bscroll = True
    '100304 rjk: Protect vsTech Min/Max
    If Val(txtTech.Text) > vsTech.Min Then
        vsTech.Value = vsTech.Min
        txtTech.Text = CStr(vsTech.Value)
    ElseIf Val(txtTech.Text) < vsTech.Max Then
        vsTech.Value = vsTech.Max
        txtTech.Text = CStr(vsTech.Value)
    Else
        vsTech.Value = Val(txtTech.Text)
    End If
    bscroll = False
End If

End Sub

Private Sub vsTech_Change()
If Not bscroll Then
    bscroll = True
    txtTech.Text = CStr(vsTech.Value)
    bscroll = False
End If
CalculateRanges CSng(vsTech.Value)
End Sub

Private Sub CalculateRanges(Tech As Single)
If frmSectors.Visible Then
    CalculateSectorRanges Tech
ElseIf frmShip.Visible Then
    CalculateShipRanges Tech
ElseIf frmLand.Visible Then
    CalculateLandRanges Tech
End If
If frmNuke.Visible Then
    If TabStrip1.SelectedItem.Index = 4 Then
        CalculateNukeDamage
    Else
        CalculatePlaneRange vsTech.Value
    End If
Else
    If Len(NukeDamageType) > 0 Or Len(PlaneRangeType) > 0 Then
        NukeDamageType = vbNullString
        PlaneRangeType = vbNullString
        frmDrawMap.DrawHexes
    End If
End If
End Sub

Private Sub CalculateSectorRanges(Tech As Single)
'Sector Values
lblFort100 = Format$(FortRange(Tech), "#0.00")
lblFort59 = Format$(FortRange(Tech, 59), "#0.00")
lblRadar = CStr(RadarRange(Tech))
lblCoastwatch = CStr(CoastwatchRange(Tech))
End Sub

Private Sub CalculateShipRanges(Tech As Single)
Dim frnge As Integer
Dim speed As Integer
Dim movecost As Single
If cmbShip.ListIndex < 0 Then Exit Sub

If Tech >= cmbShip.ItemData(cmbShip.ListIndex) Then
    Label4.Caption = "Firing Range"
    frnge = ComputeRangeStat(Tech, "s", Trim$(Right$(cmbShip.List(cmbShip.ListIndex), 5)))
    lblShipRange.Caption = Format$(UnitFireRange(Tech, frnge), "#0.00")
    speed = ComputeSpeedStat(Tech, "s", Trim$(Right$(cmbShip.List(cmbShip.ListIndex), 5)))
    If speed <= 0 Then
        lblShipMoveCost.Caption = vbNullString
        lblShipTurn.Caption = vbNullString
        lblShipMax.Caption = vbNullString
    Else
        movecost = ShipMoveCost(speed, 100, CInt(Tech), _
            Trim$(Right$(cmbShip.List(cmbShip.ListIndex), 5)))
        lblShipMoveCost.Caption = Format$(movecost, "#0.0")
        lblShipTurn.Caption = Format$((ShipMobTurn / movecost), "#0.0")
        lblShipMax.Caption = Format$((ShipMobMax / movecost), "#0.0")
    End If
Else
    Label4.Caption = "Build Tech"
    lblShipRange.Caption = CStr(cmbShip.ItemData(cmbShip.ListIndex))
    lblShipMoveCost.Caption = vbNullString
    lblShipTurn.Caption = vbNullString
    lblShipMax.Caption = vbNullString
End If
End Sub
Private Sub CalculateLandRanges(Tech As Single)
Dim frnge As Integer
If cmbLand.ListIndex < 0 Then Exit Sub

If Tech >= cmbLand.ItemData(cmbLand.ListIndex) Then
    Label7.Caption = "Firing Range"
    frnge = ComputeRangeStat(Tech, "l", Trim$(Right$(cmbLand.List(cmbLand.ListIndex), 5)))
    lblLandRange.Caption = Format$(UnitFireRange(Tech, frnge), "#0.00")
Else
    Label7.Caption = "Build Tech"
    lblLandRange.Caption = CStr(cmbLand.ItemData(cmbLand.ListIndex))
End If
End Sub

Private Sub CalculateNukeDamage()
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, txtOrigin.Text) And Len(cmbNuke.Text) > 0 Then
    If NukeDamageType <> cmbNuke.Text Or _
       NukeTargetSector <> txtOrigin.Text Or _
       (optAirBurst.Value <> vbUnchecked) <> bNukeAirBurst Then
        NukeDamageType = cmbNuke.Text
        NukeTargetSector = txtOrigin.Text
        If optAirBurst.Value <> vbUnchecked Then
            bNukeAirBurst = True
        Else
            bNukeAirBurst = False
        End If
        frmDrawMap.DrawHexes
    End If
Else
    If Len(NukeDamageType) > 0 Then
        NukeDamageType = vbNullString
        frmDrawMap.DrawHexes
    End If
End If
End Sub

Private Sub CalculatePlaneRange(Tech As Integer)
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, txtOrigin.Text) And Len(cmbNuke.Text) > 0 Then
    If PlaneRangeType <> cmbNuke.Text Or _
       PlaneTargetSector <> txtOrigin.Text Or _
       PlaneRoundTrip <> (chkRoundTrip.Value <> vbUnchecked) Or _
       PlaneTech <> Tech Then
        PlaneRangeType = cmbNuke.Text
        PlaneTargetSector = txtOrigin.Text
        PlaneRoundTrip = chkRoundTrip.Value <> vbUnchecked
        PlaneTech = Tech
        frmDrawMap.DrawHexes
    End If
Else
    If Len(PlaneRangeType) > 0 Then
        PlaneRangeType = vbNullString
        frmDrawMap.DrawHexes
    End If
End If
End Sub

Private Sub LoadComboBox(cmb As Control, Class As String)
cmb.Clear
Dim hldIndex As String

If Class = "n" Then
    hldIndex = rsBuild.Index
    rsBuild.Index = "Tech"
    If rsBuild.BOF And rsBuild.EOF Then
        Exit Sub
    End If
    rsBuild.MoveFirst
End If

If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = Class Then
            If Class = "n" Then
                cmb.AddItem rsBuild.Fields("id") + " - " + _
                    rsBuild.Fields("desc")
            Else
                cmb.AddItem rsBuild.Fields("desc") + _
                    "                                      " _
                + Left$(rsBuild.Fields("id") + "     ", 5)
            End If
            cmb.ItemData(cmb.NewIndex) = rsBuild.Fields("tech")
        End If
        rsBuild.MoveNext
    Wend
End If
If Class = "n" Then
    rsBuild.Index = hldIndex
End If
End Sub




