VERSION 5.00
Begin VB.Form frmPromptRecon 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2040
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisplayPath 
      Caption         =   "Display Path"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   1680
      Width           =   1335
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
      Left            =   7200
      TabIndex        =   19
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cbCom 
      Height          =   315
      ItemData        =   "frmPromptRecon.frx":0000
      Left            =   5160
      List            =   "frmPromptRecon.frx":0016
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtEscort 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Takeoff !"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat "
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPromptRecon.frx":002F
      Left            =   6120
      List            =   "frmPromptRecon.frx":0042
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Mobility Cost:"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Plane Range:"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Round Trip Distance:"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "commodity"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Assembly"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   4080
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "escorts"
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
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Must Fill"
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
      TabIndex        =   10
      Top             =   120
      Width           =   855
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
      Left            =   6960
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
'            changes to accomodate how the tabs are controlled in
'            SectorPanel
'08xx03 efj: remvoed dead variables
'091803 rjk: Added Unit grid selection when activating this form or selecting fields
'            and when disactivating return to standard Sector display.
'            General reformatting. Added setting the initial field on start up.
'100203 rjk: Repeat is not appropriate for fly so it is removed from the form.
'101403 rjk: Added MT_DROP to commands requesting bmap to catch mining from a plane.
'101703 rjk: Added Strength fields to Sector database
'111803 rjk: Added DisplayPath for display path, does not do sector to sector yet.
'112003 rjk: Added option to control strength updates
'            Added red highlighting when one or more planes exceed their range
'130604 rjk: Added the ability to display path from sector to sector
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace sdump, pdump and lost with GetShips, GetLandUnits and GetLost.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'210506 rjk: Fixed combobox commodities for SP: Atlantis

Dim MissionType As Integer
Dim PlaneType As enumUnitType
Public ShowRange As Boolean

Const MT_FLY = 1
Const MT_RECON = 2
Const MT_DROP = 3
Const MT_PARADROP = 4

Private Sub chkDisplayPath_Click()
SetMobilityCost
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Private Sub Combo1_Change()
Check1.Value = vbChecked
End Sub

Private Sub Combo1_Click()
Check1.Value = vbChecked
End Sub

Private Sub Label2_Change()
Select Case LCase$(Trim$(Label2.Caption))
Case "recon"
    MissionType = MT_RECON
    PlaneType = TYPE_P_SPY
    cbCom.Visible = False
    Label7.Visible = False
Case "sweep"
    MissionType = MT_RECON
    PlaneType = TYPE_P_SWEEP
    cbCom.Visible = False
    Label7.Visible = False
Case "drop"
    MissionType = MT_DROP
    PlaneType = TYPE_PC_DROP
    cbCom.Visible = True
    Label7.Visible = True
    LoadCommodityBox cbCom
    cbCom.AddItem "none"
    If frmOptions.bSPAtlantis Then
        cbCom.ListIndex = 1
    Else
        cbCom.ListIndex = 8  'set to military
    End If
Case "paradrop"
    MissionType = MT_PARADROP
    PlaneType = TYPE_P_PARA
    cbCom.Visible = False
    Label7.Visible = False
Case "fly"
    MissionType = MT_FLY
    PlaneType = TYPE_ALL
    cbCom.Visible = True
    Label7.Visible = True
    LoadCommodityBox cbCom
    cbCom.AddItem "none"
    cbCom.ListIndex = cbCom.NewIndex  'set to nothing
    Check1.Visible = False '100203 rjk: Repeat is not appropriate for fly
    Combo1.Visible = False '100203 rjk: Repeat is not appropriate for fly
    Label1.Visible = False '100203 rjk: Repeat is not appropriate for fly
End Select
End Sub

Private Sub txtEscort_GotFocus()
If Len(txtPath.Text) > 0 And ShowRange Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PG_ESCORTS, True, True, True, txtPath.Text
Else
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PG_ESCORTS, True, True, False
End If
End Sub

Private Sub txtOrigin_GotFocus()
frmDrawMap.DisplayFirstSectorPanel 'rjk 5/13/03: replace DisplaySectorPanel 1 with DisplayFirstSectorPanel
                                   '             changes to accomodate how the tabs are controlled in
                                   '             SectorPanel
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
frmPromptPath.Caption = "Select " + Label2.Caption + " Route"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless

End Sub

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
' Dim strCmd2 As String    8/12/03 efj  removed
Dim num As Integer
Dim strAP As String
Dim X As Integer
Dim n As Integer
Dim hldIndex As String
Dim CurrUnit As Variant

strCmd = LCase$(Trim$(Label2.Caption)) + " "

strCmd = strCmd + txtUnit.Text + " "

If Len(txtEscort.Text) > 0 Then
    strCmd = strCmd + txtEscort
Else
    strCmd = strCmd + "."
End If

'Verify assembly point
Dim sx As Integer
Dim sy As Integer

On Error GoTo AP_Error
strAP = txtOrigin.Text
If Not ParseSectors(sx, sy, txtOrigin.Text) Then
    'find the starting position of the first plane in the bombing run
    Dim ap As Variant
    ap = ComputeAirports(txtUnit.Text)
    If IsArray(ap) Then
        If ParseSectors(sx, sy, CStr(ap(1))) Then
            strAP = CStr(ap(1))
        Else
            GoTo AP_Error
        End If
    Else
AP_Error:
        MsgBox "Cannot compute assembly point.  Manually enter and resubmit"
        Exit Sub
    End If
End If
On Error GoTo 0

strCmd = strCmd + " " + strAP + " " + txtPath.Text

If cbCom.Visible Then
    Dim strItem As String
    strItem = cbCom.Text
    If Left$(strItem, 1) = "u" Then
        strItem = "un"
    End If
    If Left$(strItem, 4) = "none" Then
        strItem = "."
    End If
    strCmd = strCmd + " " + strItem
End If

'Execute command
If Check1.Value = vbChecked Then
    num = Val(Combo1.Text)
Else
    num = 1
End If

If Not cbCom.Visible Then
    If num > 1 Then
        frmEmpCmd.SubmitEmpireCommand "bf1", False
        For X = 1 To num
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        Next
        frmEmpCmd.SubmitEmpireCommand "bf2", False
    Else
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
Else
    frmEmpCmd.SubmitEmpireCommand "ba1 " + strItem, False
    For X = 1 To num
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    Next
    frmEmpCmd.SubmitEmpireCommand "ba2", False
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
GetPlanes
GetLost

'Check for reconissence
If MissionType = MT_RECON Or MissionType = MT_DROP Then '101403 rjk: Added MT_DROP to catch mining from a plane.
    frmEmpCmd.SubmitEmpireCommand "bmap *", False
End If

'Check to see if any planes were on ships.  If so, dump them too
CurrUnit = GetUnitList(txtUnit.Text, "p")
If IsArray(CurrUnit) Then
    hldIndex = rsPlane.Index
    rsPlane.Index = "PrimaryKey"
    n = LBound(CurrUnit) + 1
    Do While n <= UBound(CurrUnit)
        rsPlane.Seek "=", CInt(CurrUnit(n))
        If Not rsPlane.NoMatch Then
            If rsPlane.Fields("ship") > -1 Then
                frmEmpCmd.CarrierDumpRequested = True
                GetShips
                Exit Do
            End If
        End If
        n = n + 1
    Loop
    rsPlane.Index = hldIndex

    'Check to see if any of the planes were missioned
    n = LBound(CurrUnit) + 1
    Do While n <= UBound(CurrUnit)
        rsMissions.Seek "=", "p", CInt(CurrUnit(n))
        If Not rsMissions.NoMatch Then
            rsMissions.Delete
        End If
        n = n + 1
    Loop
End If

'Check if escorts moved have missions.  If so, remove them
If Len(txtEscort.Text) > 0 Then
    CurrUnit = GetUnitList(txtUnit.Text, "p")
    If IsArray(CurrUnit) Then
        n = LBound(CurrUnit) + 1
        Do While n <= UBound(CurrUnit)
            rsMissions.Seek "=", "p", CInt(CurrUnit(n))
            If Not rsMissions.NoMatch Then
                rsMissions.Delete
            End If
            n = n + 1
        Loop
    End If
End If

frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long    8/12/03 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
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
'If ShowRange Then
'    If LastLoad = 1 Then
'        LastLoad = 0
'        txtUnit_GotFocus
'    ElseIf LastLoad = 2 Then
'        LastLoad = 0
'        txtEscort_GotFocus
'    End If
'End If

If Len(txtUnit.Text) = 0 Or Len(txtPath.Text) = 0 Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
SetMobilityCost
End Sub

Private Sub txtPath_GotFocus()
If Not ShowRange Then
    frmDrawMap.DisplayFirstSectorPanel 'rjk 5/13/03: replace DisplaySectorPanel 1 with DisplayFirstSectorPanel
                                       '             changes to accomodate how the tabs are controlled in
                                       '             SectorPanel
End If
End Sub

Private Sub txtUnit_Change()
If Len(txtUnit.Text) = 0 Or Len(txtPath.Text) = 0 Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
SetMobilityCost
End Sub

Private Sub txtUnit_GotFocus()
If Len(txtPath.Text) > 0 And ShowRange Then
    frmDrawMap.SetUnitDisplay udPLANE, PlaneType, True, True, True, txtPath.Text
Else
    frmDrawMap.SetUnitDisplay udPLANE, PlaneType, True, True, False
End If
End Sub

Private Sub SetMobilityCost()
Dim sx As Integer
Dim sy As Integer
Dim sx2 As Integer
Dim sy2 As Integer
Dim mstr As String
Dim rstr As String
Dim dstr As String
Dim mcost As Single
Dim hindex As Variant
Dim distance As Integer
Dim n As Integer

On Error GoTo SetMobilityCostEnd

If MissionType = MT_DROP Or MissionType = MT_PARADROP Then
    dstr = "Round Trip Distance: "
Else
    dstr = "Trip Distance: "
End If

mstr = "Mobility Cost: "
rstr = "Plane Range: "
hindex = rsPlane.Index
rsPlane.Index = "PrimaryKey"

'must have an assembly point
If Not ParseSectors(sx2, sy2, txtOrigin.Text) Then
    GoTo SetMobilityCostEnd
End If

'must have a bomber and targets
If Len(txtUnit.Text) = 0 Then
    GoTo SetMobilityCostEnd
End If

'display path 111803 rjk: Added
If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If
If chkDisplayPath.Value = vbChecked And Len(txtPath.Text) > 0 Then
    If Not ParseSectors(sx, sy, txtPath.Text) Then
        frmDrawMap.DisplayPath txtOrigin.Text, txtPath.Text
    Else
        Dim bpath As String
        Dim pvar As Variant

        pvar = BestPath(SectorString(sx2, sy2), txtPath.Text, MT_PLANE)
        bpath = pvar(1)
        frmDrawMap.DisplayPath txtOrigin.Text, bpath
    End If
End If

'compute distance
If Len(txtPath.Text) > 0 Then
    If ParseSectors(sx, sy, txtPath.Text) Then
        distance = SectorDistance(sx, sy, sx2, sy2)
    Else
        distance = Len(txtPath.Text) - 1
    End If
Else
    distance = 0
End If

If MissionType = MT_DROP Or MissionType = MT_PARADROP Then
    distance = distance * 2
End If
dstr = dstr + CStr(distance)

'get unit list and compute mobility
Dim us As Variant
Dim c As Integer
Dim bExceededRange As Boolean '112003 rjk: Added red highlighting when one or more planes exceed their range

bExceededRange = False

c = 0
us = ParseUnitString(txtUnit.Text)
If IsArray(us) Then
    For n = LBound(us) To UBound(us)
        rsPlane.Seek "=", CStr(us(n))
        If Not rsPlane.NoMatch Then
            c = c + 1
            If c < 5 Then
                mcost = PlaneFlyCost(rsPlane.Fields("range"), rsPlane.Fields("eff"), distance)
                mstr = mstr + CStr(CInt(mcost)) + "/"
                rstr = rstr + CStr(CInt(rsPlane.Fields("range"))) + "/"
            ElseIf c = 5 Then
                mstr = mstr + "... "
                rstr = rstr + "... "
            End If
            If distance > rsPlane.Fields("range") Then '112003 rjk: Added red highlighting when one or more planes exceed their range
                bExceededRange = True
            End If
        End If
    Next n
End If


SetMobilityCostEnd:
Label8.Caption = Left$(mstr, Len(mstr) - 1)
If bExceededRange Then '112003 rjk: Added red highlighting when one or more planes exceed their range
    Label9.ForeColor = vbRed
Else
    Label9.ForeColor = vbBlack
End If
Label9.Caption = Left$(rstr, Len(rstr) - 1)
Label10.Caption = dstr
rsPlane.Index = hindex
End Sub
