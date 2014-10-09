VERSION 5.00
Begin VB.Form frmPromptNav 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2376
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   8736
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2376
   ScaleWidth      =   8736
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   975
      Left            =   3840
      TabIndex        =   16
      Top             =   1320
      Width           =   4455
      Begin VB.CheckBox chkView 
         Caption         =   "View along the way"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkRefreshBMapOnStop 
         Caption         =   "Refresh Bmap on stop"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   670
         Width           =   2175
      End
      Begin VB.CheckBox chkDisplayPath 
         Caption         =   "Display Path"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   670
         Width           =   1335
      End
      Begin VB.CheckBox chkLookAlongTheWay 
         Caption         =   "Look along the way"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   130
         Width           =   1815
      End
      Begin VB.CheckBox chkStopAfterMove 
         Caption         =   "Stop after move"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   400
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Post Move Settings"
      Height          =   1095
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtRadius 
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Text            =   "99"
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkMission 
         Caption         =   "Interdiction/Reserve Mission"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRetreatPath 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkRetreat 
         Caption         =   "Retreat on"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "radius"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "path"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblRouteInfo 
      Caption         =   "RouteInfo"
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblBestPath 
      Caption         =   "best path:"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblPathCost 
      Caption         =   "path cost:"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblCost 
      Caption         =   "mob cost:"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblLeft 
      Caption         =   "mob left:"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "path/destination"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Navigate"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
'                       replaced DisplaySectorPanel 3 with DisplayUnitBox
'                       changes to accomodate how the tabs are controlled in
'                       SectorPanel
'081203 efj: removed dead variables
'092703 rjk: general reformatting, removed hard coding for unit types
'            set OK to be the default button
'093003 rjk: added RouteInfo for indicating if multiple origins are present
'093003 rjk: if multiple origins are present can not selected a target sector.
'111803 rjk: Moved DisplayPath to frmDrawMap
'120303 rjk: Switched to frmOptions and boolean options.
'120703 rjk: Changed hldindex to hldIndex.
'220404 rjk: Code cleanup.
'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
'            ships with SWEEP capabilities.  Removed correction factor from NavMarkers.
'160705 rjk: Added ability to get OContent for Sea Sectors
'180206 rjk: Replace sdump, ldump, and pdump with GetShips, GetLandUnits and GetPlanes.
'190606 rjk: Incorporate the sector mobility changes for 4.3.6 servers.
'160706 rjk: Added AutoView for navigate and march.
'181006 rjk: Remove AutoView for march.
'140708 rjk: Added GetLost to navigate and march.

Public strSector As String
Dim UnitList As Variant

'drk 7/13/02
Private Sub chkRefreshBMapOnStop_Click()
frmNavCtrl.RefreshBmapOnStop = chkRefreshBMapOnStop.Value
End Sub

Private Sub chkDisplayPath_Click()
ComputePathCost
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim strCmd2 As String
Dim wp As Variant
Dim X As Integer
Dim n As Integer
' Dim rs As Recordset    8/12/03 efj  removed
Dim strType As String
Dim CurrUnit As Variant
Dim strInterdict As String
Dim strReserve As String
Dim hldIndex As Variant


If Len(txtUnit.Text) = 0 Then
    Exit Sub
End If

frmNavCtrl.AutoLooking_when_notAutoStopping = False
frmNavCtrl.AutoViewing_when_notAutoStopping = False

'if the path is blank, ignore the autostop
strCmd = LCase$(Label2.Caption) + " " + txtUnit.Text + " "
strCmd2 = vbNullString

wp = ParseWaypoints(txtPath.Text)
If IsArray(wp) Then
    frmEmpCmd.SubmitEmpireCommand "bf1", True
    For X = 1 To UBound(wp)
        If chkLookAlongTheWay.Value = vbChecked Then
            If chkView.Value = vbChecked And chkView.Visible Then
                frmEmpCmd.SubmitEmpireCommand "as1lv", False
            Else
                frmEmpCmd.SubmitEmpireCommand "as1l", False
            End If
        ElseIf chkView.Value = vbChecked And chkView.Visible Then
            frmEmpCmd.SubmitEmpireCommand "as1v", False
        Else
            frmEmpCmd.SubmitEmpireCommand "as1 ", False
        End If
        strCmd2 = strCmd + " " + wp(X)
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    Next X
    frmEmpCmd.SubmitEmpireCommand "bf2", True
Else
    'Submit the autostop
    If chkStopAfterMove.Value = vbChecked And Len(txtPath.Text) > 0 Then
        frmNavCtrl.AutoLooking_when_notAutoStopping = False
        frmNavCtrl.AutoViewing_when_notAutoStopping = False
        If chkLookAlongTheWay.Value = vbChecked Then
            If chkView.Value = vbChecked And chkView.Visible Then
                frmEmpCmd.SubmitEmpireCommand "as1lv", False
            Else
                frmEmpCmd.SubmitEmpireCommand "as1l", False
            End If
        ElseIf chkView.Value = vbChecked And chkView.Visible Then
            frmEmpCmd.SubmitEmpireCommand "as1v", False
        Else
            frmEmpCmd.SubmitEmpireCommand "as1 ", False
        End If
    ElseIf chkLookAlongTheWay.Value = vbChecked Or (chkView = vbChecked And chkView.Visible) Then
        If chkLookAlongTheWay.Value Then
            'no autostop, but we do want to look along the way
            frmNavCtrl.AutoLooking_when_notAutoStopping = True
        End If
        If chkView.Value = vbChecked And chkView.Visible Then
            frmNavCtrl.AutoViewing_when_notAutoStopping = True
        End If
    End If
    
    strCmd = strCmd + txtPath
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If

'Check if units moved have missions.  If so, remove them
If LCase$(Label2.Caption) = "march" Then
    strType = "l"
Else
    strType = "s"
End If
CurrUnit = GetUnitList(txtUnit.Text, strType)
If IsArray(CurrUnit) Then

    n = LBound(CurrUnit) + 1
    Do While n <= UBound(CurrUnit)
        rsMissions.Seek "=", strType, CInt(CurrUnit(n))
        If Not rsMissions.NoMatch Then
            rsMissions.Delete
        End If
        n = n + 1
    Loop
End If

'set retreat
frmDrawMap.DoRetreat = False
If chkStopAfterMove.Value = vbChecked And Len(txtPath.Text) > 0 Then
    If chkRetreat.Value = vbChecked And Len(txtRetreatPath.Text) > 0 Then
        strCmd = "retreat " + txtUnit.Text + " " + txtRetreatPath.Text + " " + Left$(Combo1.Text, 1)
        If LCase$(Label2.Caption) = "march" Then
            strCmd = "l" + strCmd
        End If
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
ElseIf chkRetreat.Value <> vbUnchecked Then
    'Set up retreat string for NavCtrl prompt
    frmDrawMap.DoRetreat = True
    frmDrawMap.LastRetreatString = txtRetreatPath.Text
    frmDrawMap.LastRetreatUnits = txtUnit.Text
    If LCase$(Label2.Caption) = "march" Then
        frmDrawMap.LastRetreatType = "l"
    Else
        frmDrawMap.LastRetreatType = vbNullString
    End If
End If

'Set post move missions if requested
If chkMission.Value = vbChecked Then
    If Val(txtRadius.Text) <= 0 Then
        txtRadius.Text = vbNullString
    End If
    If strType = "s" Then
        'ships are alway interdiction
        strCmd = "mission ship " + txtUnit.Text + " i . " + txtRadius.Text
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    
        'land can be interdiction or reserve - figure out which fits
    ElseIf IsArray(CurrUnit) Then
        n = LBound(CurrUnit) + 1
        strInterdict = vbNullString
        strReserve = vbNullString
        hldIndex = rsLand.Index
        rsLand.Index = "PrimaryKey"
        Do While n <= UBound(CurrUnit)
            rsLand.Seek "=", CInt(CurrUnit(n))
            If Not rsLand.NoMatch Then
                If rsLand.Fields("frg") > 0 Then
                    strInterdict = strInterdict + CStr(CurrUnit(n)) + "/"
                ElseIf rsLand.Fields("def") >= 1 Then
                    strReserve = strReserve + CStr(CurrUnit(n)) + "/"
                End If
            End If
            n = n + 1
        Loop
        hldIndex = rsLand.Index
        If Len(strInterdict) > 0 Then
            strCmd = "mission land " + Left$(strInterdict, Len(strInterdict) - 1) + " i . " + txtRadius.Text
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
        If Len(strReserve) > 0 Then
            strCmd = "mission land " + Left$(strReserve, Len(strReserve) - 1) + " r . " + txtRadius.Text
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
    End If
End If
        
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If LCase$(Label2.Caption) = "navigate" Then
    DoNavDumps
    GetOContent
ElseIf LCase$(Label2.Caption) = "march" Then
    GetLandUnits
Else '093003 rjk: This case should not happen as only nav and march use this form.
    GetPlanes
End If
GetLost
        

cmdOK_exit:
frmEmpCmd.SubmitEmpireCommand "db2", False

If chkStopAfterMove.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptNav", "StopAfterMove", "true"
Else
    SaveSetting APPNAME, "frmPromptNav", "StopAfterMove", "false"
End If
If chkLookAlongTheWay.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptNav", "LookAlongTheWay", "true"
Else
    SaveSetting APPNAME, "frmPromptNav", "LookAlongTheWay", "false"
End If
If chkRefreshBMapOnStop.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptNav", "RefreshBMapOnStop", "true"
Else
    SaveSetting APPNAME, "frmPromptNav", "RefreshBMapOnStop", "false"
End If
If chkDisplayPath.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptNav", "DisplayPath", "true"
Else
    SaveSetting APPNAME, "frmPromptNav", "DisplayPath", "false"
End If
If LCase$(Label2.Caption) = "navigate" Then
    If chkView.Value <> vbUnchecked Then
        SaveSetting APPNAME, "frmPromptNav", "View", "true"
    Else
        SaveSetting APPNAME, "frmPromptNav", "View", "false"
    End If
End If
Unload Me
End Sub

Private Sub Form_Activate()
If LCase$(Label2.Caption) = "navigate" Then
    chkView.Visible = True
Else
    chkView.Visible = False
End If
End Sub

Private Sub Form_Load()
'Make Stay always on top
'Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
ComputePathCost
LoadRetreatBox Combo1
Combo1.ListIndex = 0
If frmOptions.bDefaultRetreat Then
    chkRetreat.Value = vbChecked
Else
    chkRetreat.Value = vbUnchecked
End If

'drk: persist these settings through the registry. I'd really prefer a
'different method, but this works for now
If GetSetting(APPNAME, "frmPromptNav", "StopAfterMove", "false") = "false" Then
    chkStopAfterMove.Value = vbUnchecked
Else
    chkStopAfterMove.Value = vbChecked
End If
If GetSetting(APPNAME, "frmPromptNav", "LookAlongTheWay", "false") = "false" Then
    chkLookAlongTheWay.Value = vbUnchecked
Else
    chkLookAlongTheWay.Value = vbChecked
End If
If GetSetting(APPNAME, "frmPromptNav", "RefreshBMapOnStop", "false") = "false" Then
    chkRefreshBMapOnStop.Value = vbUnchecked
Else
    chkRefreshBMapOnStop.Value = vbChecked
End If
If GetSetting(APPNAME, "frmPromptNav", "DisplayPath", "false") = "false" Then
    chkDisplayPath.Value = vbUnchecked
Else
    chkDisplayPath.Value = vbChecked
End If
If GetSetting(APPNAME, "frmPromptNav", "View", "false") = "false" Then
    chkView.Value = vbUnchecked
Else
    chkView.Value = vbChecked
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If
End Sub

Private Sub txtPath_Change()
ComputePathCost
End Sub

Private Sub txtPath_DblClick()
'Make sure starting sector is valid before prompting
Dim sx As Integer
Dim sy As Integer
If Not ParseSectors(sx, sy, strSector) Then
    Exit Sub
End If

Load frmPromptPath
frmPromptPath.lblSector.Caption = strSector
frmPromptPath.lblEndSector.Caption = strSector
frmPromptPath.Caption = "Select Route"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless
End Sub

Private Sub txtPath_GotFocus()
frmDrawMap.DisplayFirstSectorPanel 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
                                   '             changes to accomodate how the tabs are controlled in
                                   '             SectorPanel
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    KeyAscii = 0
    Load frmPromptWaypoints
    frmPromptWaypoints.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptWaypoints.Width
    frmPromptWaypoints.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptWaypoints.Height) \ 2
    frmPromptWaypoints.Show vbModeless
    frmPromptWaypoints.txtOrigin.SetFocus

End If
End Sub

Private Sub DoNavDumps()
Dim sx As Integer
Dim sy As Integer
Dim nID As Integer
Dim hldIndex1 As Variant
Dim hldIndex2 As Variant
Dim hldIndex3 As Variant
Dim LandFound As Boolean
Dim PlaneFound As Boolean
Dim oldx As Integer
Dim oldy As Integer
Dim n As Integer

GetShips
    
If Not IsArray(UnitList) Then
    Exit Sub
End If
    
hldIndex1 = rsShip.Index
hldIndex2 = rsPlane.Index
hldIndex3 = rsLand.Index
rsShip.Index = "PrimaryKey"
rsPlane.Index = "location"
rsLand.Index = "location"
    
'if there are planes and units in the save starting sector as
'the ships, then go ahead and do the ldump and pdump as well
n = 1
While n <= UBound(UnitList) And Not (LandFound And PlaneFound)
    LandFound = False
    PlaneFound = False
    oldx = -1
    oldy = 0
    nID = CInt(UnitList(n))
    rsShip.Seek "=", nID
    If Not rsShip.NoMatch Then
        sx = rsShip.Fields("x").Value
        sy = rsShip.Fields("y").Value
        If (oldx <> sx) Or (oldy <> sy) Then
            'check for land units in the sector
            If Not LandFound Then
                rsLand.Seek "=", sx, sy
                LandFound = Not rsLand.NoMatch
            End If
            'check for planes in the sector
            If Not PlaneFound Then
                rsPlane.Seek "=", sx, sy
                PlaneFound = Not rsPlane.NoMatch
            End If
            oldx = sx
            oldy = sy
        End If
    End If
    n = n + 1
Wend

If PlaneFound Then
    GetPlanes
End If

If LandFound Then
    GetLandUnits
End If
    
'restore indexes
rsShip.Index = hldIndex1
rsPlane.Index = hldIndex2
rsLand.Index = hldIndex3
End Sub

Private Sub ComputePathCost()
On Error Resume Next
Dim pvar As Variant
Dim pcost As Single
Dim mcost As Single
Dim mused As Single
Dim mleft As Single
Dim minmob As Single
Dim oldx As Integer
Dim oldy As Integer
Dim sx As Integer
Dim sy As Integer
Dim n As Integer
' Dim m As Integer    8/12/03 efj  removed
Dim nx As Integer
Dim ny As Integer
' Dim pb As Integer    8/12/03 efj  removed
Dim hldIndex As Variant
Dim rs As Recordset
Dim Fleet As String
Dim bpath As String
Dim strRetreat As String
Dim MobType As Integer

lblLeft.Caption = vbNullString
lblCost.Caption = vbNullString
lblBestPath.Caption = vbNullString
lblPathCost.Caption = vbNullString
'093003 rjk: added RouteInfo for indicating if multiple origins are present
lblRouteInfo.Caption = vbNullString

If LCase$(Label2.Caption) <> "march" And LCase$(Label2.Caption) <> "navigate" Then
    Exit Sub
End If

If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If

If Not IsArray(UnitList) Then
    Exit Sub
End If

If Len(txtPath.Text) = 0 Then
    Exit Sub
End If

'set indexs
If LCase$(Label2.Caption) = "march" Then
    Set rs = rsLand
    Fleet = "army"
    MobType = MT_LAND
Else
    Set rs = rsShip
    Fleet = "flt"
    MobType = MT_SHIP
End If
hldIndex = rs.Index
rs.Index = "PrimaryKey"
minmob = 9999
oldx = -1
oldy = 0

'093003 rjk: if multiple origins are present can not selected a target sector.
cmdOK.Enabled = True

For n = LBound(UnitList) + 1 To UBound(UnitList)
    rs.Seek "=", CInt(UnitList(n))
    If Not rs.NoMatch Then
        If MobType = MT_LAND Then
            If frmDrawMap.UnitCharacteristicCheck(TYPE_L_TRAIN, rs.Fields("type")) Then
                MobType = MT_RAIL
            End If
        ElseIf MobType = MT_SHIP Then
            If frmDrawMap.UnitCharacteristicCheck(TYPE_S_CANAL, rs.Fields("type")) Then
                MobType = MT_SMALL_SHIP
            End If
        End If
        sx = rs.Fields("x").Value
        sy = rs.Fields("y").Value
        '093003 rjk: added RouteInfo for indicating if multiple origins are present
        If (oldx = -1) And (oldy = 0) Then 'Not first location.
            '093003 rjk: Used first unit for costs and paths
            'Compute the path cost
            oldx = sx
            oldy = sy
            If ParseSectors(nx, ny, txtPath.Text) Then
                pvar = BestPath(SectorString(sx, sy), txtPath.Text, MobType)
                bpath = pvar(1)
                pcost = pvar(2)
            Else
                pcost = PathCost(SectorString(sx, sy), txtPath.Text, MobType)
                bpath = txtPath.Text
            End If
        ElseIf (sx <> oldx) Or (sy <> oldy) Then
            If ParseSectors(nx, ny, txtPath.Text) Then
                lblRouteInfo.ForeColor = vbRed
                lblRouteInfo.Caption = "Error: Multiple Origins"
                '093003 rjk: if multiple origins are present can not selected a target sector.
                cmdOK.Enabled = False
            Else
                lblRouteInfo.ForeColor = vbBlack
                lblRouteInfo.Caption = "Warning: Multiple Origins"
            End If
        'Else -same location do nothing
        End If
        
        'Fill in retreat path
        txtRetreatPath.Text = vbNullString
        strRetreat = ReversePath(bpath)
        If Len(strRetreat) > 5 Then
            strRetreat = Left$(strRetreat, 5)
        End If
        txtRetreatPath.Text = strRetreat
       
        If Fleet = "flt" Then
            mcost = ShipMoveCost(rs.Fields("spd"), rs.Fields("eff"), rs.Fields("tech"), _
                rs.Fields("type"))
            mused = mcost * pcost
        Else
            mcost = ShipMoveCost(rs.Fields("spd"), 100, rs.Fields("tech"), rs.Fields("type"))
            
            'only supply units and train suffer mob penalties for low eff
            '092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
            'If rs.Fields("eff") < 100 And (rs.Fields("type") = "sup" Or rs.Fields("type") = "tra") Then
            If rs.Fields("eff") < 100 _
               And frmDrawMap.UnitCharacteristicCheck(TYPE_L_SUPPLY, rs.Fields("type")) _
                     Then
                mcost = (mcost * 100) / rs.Fields("eff")
            End If
            mused = mcost * pcost * 5
        End If
        
        mleft = rs.Fields("mob") - mused
        If mleft < minmob Then
            minmob = mleft
            lblCost.Caption = "mob. cost: " + CStr(CSng(CLng(CSng(mused) * 1000) / 1000))
            lblLeft.Caption = "mob. left: " + CStr(Int(mleft))
            lblBestPath.Caption = "path: " + bpath
            lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
            If chkDisplayPath.Value = vbChecked Then
                frmDrawMap.DisplayPath SectorString(sx, sy), bpath
            End If
        End If
    End If
Next n
rs.Index = hldIndex
End Sub

Private Sub txtUnit_Change()
If LCase$(Label2.Caption) = "march" Then
    UnitList = GetUnitList(txtUnit.Text, "l")
Else
    UnitList = GetUnitList(txtUnit.Text, "s")
End If

ComputePathCost
End Sub

Private Sub txtUnit_GotFocus()
frmDrawMap.DisplayUnitBox 'rjk 5/13/03: replace DisplaySectorPanel 3 with DisplayUnit
                          '             changes to accomodate how the tabs are controlled in
                          '             SectorPanel
End Sub
