VERSION 5.00
Begin VB.Form frmPromptBomb 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2085
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisplayPath 
      Caption         =   "Display Path"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
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
      Left            =   7080
      TabIndex        =   24
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmPromptBomb.frx":0000
      Left            =   3720
      List            =   "frmPromptBomb.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   960
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "pinpoint"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "strategic"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtEscort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Takeoff !"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat "
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPromptBomb.frx":0043
      Left            =   5880
      List            =   "frmPromptBomb.frx":0056
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Round Trip Distance:"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Plane Range:"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Mobility Cost:"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "type of run"
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
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Assembly Point"
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
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "target/flight path"
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
      TabIndex        =   17
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
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
      Left            =   840
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Bomb"
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
      TabIndex        =   14
      Top             =   240
      Width           =   615
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
      Left            =   6720
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptBomb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
'            changes to accomodate how the tabs are controlled in SectorPanel
'092803 rjk: switched to SetUnitType and DisplayFirstSectorPanel.
'            general reformatting.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'060604 rjk: Added Display Path for Bombing
'130604 rjk: Added the ability to display path from sector to sector
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'080805 rjk: Fixed the Commodities box to start with military.
'180206 rjk: Replace ldump, sdump, lost with GetLandUnits, GetShips and GetLost.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'170506 rjk: Added infra for SP: Atlantis

Public ShowRange As Boolean

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

Private Sub Combo2_Change()
Select Case Combo2.ListIndex
Case 0
    Combo3.Visible = True
    txtTarget.Visible = False
Case 1, 2, 3
    txtTarget.Visible = True
    Combo3.Visible = False
Case Else
    txtTarget.Visible = False
    Combo3.Visible = False
End Select

End Sub

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
Case 0
    Combo3.Visible = True
    txtTarget.Visible = False
Case 1, 2, 3
    txtTarget.Visible = True
    Combo3.Visible = False
Case Else
    txtTarget.Visible = False
    Combo3.Visible = False
End Select

End Sub

Private Sub Option1_Click()
Combo2.Visible = False
Combo3.Visible = False
txtTarget.Visible = False
End Sub

Private Sub Option2_Click()
Combo2.Visible = True
Combo2.ListIndex = 0
Combo3.Visible = True
End Sub

Private Sub txtEscort_GotFocus()
If Len(txtPath.Text) > 0 And ShowRange Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PG_ESCORTS, True, True, True, txtPath.Text
Else
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PG_ESCORTS, True, True, False
End If
End Sub

Private Sub txtOrigin_GotFocus()
frmDrawMap.DisplayFirstSectorPanel 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
                                   '             changes to accomodate how the tabs are controlled in
                                   '             SectorPanel
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
frmPromptPath.Caption = "Select Bomb Route"
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
Dim strCmd2 As String
Dim strItem As String
Dim strAP As String
Dim num As Integer
Dim X As Integer
Dim n As Integer
Dim hldIndex As String
Dim CurrUnit As Variant

'Must have a destination, or will bomb yourself
If Len(txtPath.Text) = 0 Then
    MsgBox "Must have a destination/path, or you will bomb yourself !"
    Exit Sub
End If

strItem = LCase$(Combo3.Text)
If Left$(strItem, 1) = "u" Then
    strItem = "un"
End If

strCmd = LCase$(Trim$(Label2.Caption)) + " "

strCmd = strCmd + txtUnit.Text + " "

If Len(txtEscort.Text) > 0 Then
    strCmd = strCmd + txtEscort
Else
    strCmd = strCmd + "."
End If

If Option1.Value Then
    strCmd = strCmd + " " + Option1.Caption
Else
    strCmd = strCmd + " " + Option2.Caption
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

'Build bomb command
strCmd = strCmd + " " + strAP + " " + txtPath.Text

'Execute fire command
If Check1.Value = vbChecked Then
    num = Val(Combo1.Text)
Else
    num = 1
End If

If Option2.Value Then   'pinbombing
    strCmd2 = Left$(Combo2.Text, 3)
    If Combo2.ListIndex = 0 Then   'commodities
        strCmd2 = "com" + Left$(Combo3.Text, 3)
    ElseIf Combo2.ListIndex < 4 Then
        strCmd2 = strCmd2 + txtTarget.Text
    End If
End If
        
frmEmpCmd.SubmitEmpireCommand "ba1 " + strCmd2, False
For X = 1 To num
    frmEmpCmd.SubmitEmpireCommand strCmd, True
Next
frmEmpCmd.SubmitEmpireCommand "ba2", False

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False

'Dump the airports from which the planes are flying
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
   
'Check to see if any planes were on ships.  If so, dump them to
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

'now dump the planes themselves
GetPlanes
GetLost
frmEmpCmd.SubmitEmpireCommand "db2", False

If chkDisplayPath.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptBomb", "Display Path", True
Else
    SaveSetting APPNAME, "frmPromptBomb", "Display Path", False
End If

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub


Private Sub Form_Load()

' Set parent for the toolbar to display on top of:
' Dim success As Long    8/12/03 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

txtTarget.Move Combo3.Left, Combo3.top
LoadCommodityBox Combo3
Combo3.ListIndex = 1 'set to military
Combo2.ListIndex = 0

If GetSetting(APPNAME, "frmPromptBomb", "Display Path", False) Then
    chkDisplayPath.Value = vbChecked
Else
    chkDisplayPath.Value = vbUnchecked
End If

If frmOptions.bSPAtlantis Then
    Combo2.AddItem "infra"
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


Private Sub txtPath_GotFocus()
If Not ShowRange Then
    frmDrawMap.DisplayFirstSectorPanel 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
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
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PC_BOMB, True, True, True, txtPath.Text
Else
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_PC_BOMB, True, True, False
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
Dim bpath As String
Dim pvar As Variant

On Error GoTo SetMobilityCostEnd

mstr = "Mobility Cost: "
dstr = "Round Trip Distance: "
rstr = "Plane Range: "
hindex = rsPlane.Index
rsPlane.Index = "PrimaryKey"

If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If

'must have an assembly point
If Not ParseSectors(sx2, sy2, txtOrigin.Text) Then
    GoTo SetMobilityCostEnd
End If

'must have a bomber and targets
If Len(txtUnit.Text) = 0 Then
    GoTo SetMobilityCostEnd
End If

'compute distance
If Len(txtPath.Text) > 0 Then
    If ParseSectors(sx, sy, txtPath.Text) Then
        distance = SectorDistance(sx, sy, sx2, sy2)
        If chkDisplayPath.Value <> vbUnchecked Then
            pvar = BestPath(SectorString(sx2, sy2), txtPath.Text, MT_PLANE)
            bpath = pvar(1)
        End If
    Else
        If chkDisplayPath.Value <> vbUnchecked Then
            bpath = txtPath.Text
        End If
        distance = Len(txtPath.Text) - 1
    End If
Else
    distance = 0
End If
dstr = dstr + CStr(distance * 2)

'get unit list and compute mobility
Dim us As Variant
Dim c As Integer
c = 0
us = ParseUnitString(txtUnit.Text)
If IsArray(us) Then
    For n = LBound(us) To UBound(us)
        rsPlane.Seek "=", CStr(us(n))
        If Not rsPlane.NoMatch Then
            c = c + 1
            If c < 5 Then
                mcost = PlaneFlyCost(rsPlane.Fields("range"), rsPlane.Fields("eff"), 2 * distance)
                mstr = mstr + CStr(CInt(mcost)) + "/"
                rstr = rstr + CStr(CInt(rsPlane.Fields("range"))) + "/"
            ElseIf c = 5 Then
                mstr = mstr + "... "
                rstr = rstr + "... "
            End If
        End If
    Next n
End If

If chkDisplayPath.Value <> vbUnchecked Then
    frmDrawMap.DisplayPath SectorString(sx2, sy2), bpath
End If


SetMobilityCostEnd:
Label8.Caption = Left$(mstr, Len(mstr) - 1)
Label9.Caption = Left$(rstr, Len(rstr) - 1)
Label10.Caption = dstr
rsPlane.Index = hindex
End Sub

