VERSION 5.00
Begin VB.Form frmPromptTrans 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuantity 
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   735
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
      Left            =   6720
      TabIndex        =   16
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkDisplayPath 
      Caption         =   "Display Path"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1215
      Begin VB.OptionButton Option1 
         Caption         =   "plane"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "nuke"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Label lblOrigin 
      Alignment       =   2  'Center
      Caption         =   "start"
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
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblQuantity 
      Alignment       =   2  'Center
      Caption         =   "quantity"
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
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLeft 
      Caption         =   "mob left:"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblCost 
      Caption         =   "mob cost:"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblPathCost 
      Caption         =   "path cost:"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblBestPath 
      Caption         =   "best path:"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "path/ dest."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "unit"
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
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Transport"
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
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'xx0803 efj: removed dead variables
'180903 rjk: Added Unit grid selection when activating this form or selecting fields
'            and when disactivating return to standard Sector display.
'            General reformatting. Added setting the initial field on start up.
'250903 rjk: Added Quantity for nukes.
'            Added Origin label.  Added timestamp ndump. Added path cost
'            for nukes.
'021003 rjk: Added check to see if there is also a nuke or plane in the sector before enabling second option.
'051003 rjk: Removed timestamp ndump
'061003 rjk: Changed Transport submit command to True so it appears in the F9 list.
'141003 rjk: Added ndump timestamp (try #2)
'171003 rjk: Added Strength fields to Sector database
'181103 rjk: Moved DisplayPath to frmDrawMap
'201103 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace pdump, ndump and lost with GetPlanes, GetNukes and GetLost.
'120305 rjk: Fixed transport costs for SP2005.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'060506 rjk: Added supported for Nuke id changes in 4.3.3 servers.
'140506 rjk: Fixed nuke transport to select nuke grid.
'190606 rjk: Incorporate the sector mobility changes for 4.3.6 servers.
'030706 rjk: Fixed the sector updating for nukes.  Update the grid when transporting.

Public strCmd As String
Public strSector As String

Private Sub chkDisplayPath_Click()
ComputePathCost
End Sub

Private Sub cmdCancel_Click()
strCmd = vbNullString
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim X As Integer
Dim wp As Variant
Dim strFill1 As String
Dim strFill2 As String
Dim n As Integer
Dim CurrUnit As Variant

On Error Resume Next

'Start with first list box results
If Option1.Value Then
    strCmd = "plane"
    strFill1 = " "
    strFill2 = " "
Else
    strCmd = "nuke"
    If VersionCheck(4, 3, 3) <> VER_LESS Then
        strFill1 = " "
        strFill2 = " "
    Else
        strFill1 = " " + strSector + " "
        strFill2 = " " + txtQuantity.Text + " "
    End If
End If

wp = ParseWaypoints(txtPath.Text)
If IsArray(wp) Then
    Dim strCmd2 As String
    frmEmpCmd.SubmitEmpireCommand "bf1", True
    For X = 1 To UBound(wp)
        strCmd2 = "transport " + strCmd + strFill1 + txtUnit.Text + strFill2 + wp(X)
        If strCmd = "nuke" And VersionCheck(4, 3, 3) = VER_LESS Then
                strFill1 = " " + wp(X) + " "
        End If
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    Next X
    frmEmpCmd.SubmitEmpireCommand "bf2", True
Else
    strCmd = "transport " + strCmd + strFill1 + txtUnit.Text + strFill2 + txtPath
    frmEmpCmd.SubmitEmpireCommand strCmd, True '100603 rjk: Changed to True so the command appears in the F9 list.
End If

'Check if planes moved have missions.  If so, remove them
If Len(txtPath.Text) > 0 And Option1.Value Then
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

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If Option1.Value Then
    GetPlanes
End If
If Option2.Value Then
    GetNukes
    GetLost
End If
    
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Form_Activate()
'100203 rjk: Added check to see if there is also a nuke or plane in the sector before enabling second option.
Dim sx As Integer
Dim sy As Integer

lblOrigin = strSector
If Option1 Then
    txtQuantity.Visible = False
    lblQuantity.Visible = False
    If ParseSectors(sx, sy, strSector) Then
        rsNuke.Seek "=", sx, sy
        If rsNuke.NoMatch Then
            Option2.Enabled = False
        End If
    Else
        Option2.Enabled = False
    End If
Else
    If VersionCheck(4, 3, 3) <> VER_LESS Then
        txtQuantity.Visible = False
        txtQuantity.Text = "1"
        lblQuantity.Visible = False
    Else
        txtQuantity.Visible = True
        lblQuantity.Visible = True
    End If
    If ParseSectors(sx, sy, strSector) Then
        rsPlane.Seek "=", sx, sy
        If rsPlane.NoMatch Then
            Option1.Enabled = False
        End If
    Else
        Option1.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
txtPath = strSector
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False

If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If

End Sub

Private Sub Option1_Click()
ComputePathCost
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
txtQuantity.Visible = False
txtQuantity.Text = "1"
lblQuantity.Visible = False
End Sub

Private Sub Option2_Click()
ComputePathCost
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
If VersionCheck(4, 3, 3) <> VER_LESS Then
    txtQuantity.Visible = False
    txtQuantity.Text = "1"
    lblQuantity.Visible = False
Else
    txtQuantity.Visible = True
    lblQuantity.Visible = True
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
frmPromptPath.Caption = "Select Transport Route"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless
End Sub

Private Sub ComputePathCost()
On Error Resume Next
Dim pvar As Variant
Dim pcost As Single
Dim pweight As Integer
Dim mused As Single
Dim mleft As Single
Dim tweight As Long
Dim oldx As Integer
Dim oldy As Integer
Dim sx As Integer
Dim sy As Integer
Dim n As Integer
Dim nx As Integer
Dim ny As Integer
Dim hldIndex As Variant
Dim rs As Recordset
Dim bpath As String
Dim UnitList As Variant
Dim strNukeType As String
Dim quantity As Integer

lblLeft.Caption = vbNullString
lblCost.Caption = vbNullString
lblBestPath.Caption = vbNullString
lblPathCost.Caption = vbNullString

If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If

'planes only for now
' 092503 rjk: Added code for Nukes
'If Not Option1.Value Then
'    Exit Sub
'End If

If Len(txtPath.Text) = 0 Then
    Exit Sub
End If

'Don't do costs for waypoints
If InStr(txtPath.Text, ";") > 0 Then
    Exit Sub
End If

If Option1.Value Then
    UnitList = GetUnitList(txtUnit.Text, "p")
    If Not IsArray(UnitList) Then
        Exit Sub
    End If
    
    'set indexs
    Set rs = rsPlane
    hldIndex = rs.Index
    rs.Index = "PrimaryKey"
    tweight = 0
    oldx = -1
    oldy = 0
    
    For n = LBound(UnitList) To UBound(UnitList)
        If Len(UnitList(n)) > 0 Then
            rs.Seek "=", CInt(UnitList(n))
            If Not rs.NoMatch Then
                sx = rs.Fields("x").Value
                sy = rs.Fields("y").Value
                If (oldx <> sx) Or (oldy <> sy) Then
                    'Compute the path cost
                    If ParseSectors(nx, ny, txtPath.Text) Then
                        pvar = BestPath(SectorString(sx, sy), txtPath.Text, MT_COMMODITY)
                        bpath = pvar(1)
                        pcost = pvar(2)
                    Else
                        pcost = PathCost(SectorString(sx, sy), txtPath.Text, MT_COMMODITY)
                        bpath = txtPath.Text
                    End If
                    oldx = sx
                    oldy = sy
                End If
                rsBuild.Seek "=", "p", rs.Fields("type")
                If rsBuild.NoMatch Then
                    rs.Index = hldIndex
                    Exit Sub
                End If
                If frmOptions.bSP2005 Then
                    pweight = rsBuild.Fields("avail") - 20
                Else
                    pweight = rsBuild.Fields("hcm") * 2 + rsBuild.Fields("lcm")
                End If
                tweight = tweight + pweight
            End If
        End If
    Next n
Else 'Nukes
    If VersionCheck(4, 3, 6) <> VER_LESS Then
        UnitList = GetUnitList(txtUnit.Text, "n")
        If Not IsArray(UnitList) Then
            Exit Sub
        End If
        
        hldIndex = rsNuke.Index
        rsNuke.Index = "id"
        For n = LBound(UnitList) To UBound(UnitList)
            If Len(UnitList(n)) > 0 Then
                rsNuke.Seek "=", CInt(UnitList(n))
                If Not rsNuke.NoMatch Then
                    rsBuild.Seek "=", "n", rsNuke.Fields("type")
                    If Not rsBuild.NoMatch Then
                        tweight = tweight + rsBuild.Fields("stat3").Value 'stat3 is lbs field
                    End If
                    If ParseSectors(nx, ny, txtPath.Text) Then
                        pvar = BestPath(strSector, txtPath.Text, MT_COMMODITY)
                        bpath = pvar(1)
                        pcost = pvar(2)
                    Else
                        pcost = PathCost(strSector, txtPath.Text, MT_COMMODITY)
                        bpath = txtPath.Text
                    End If
                End If
            End If
        Next n
        rsNuke.Index = hldIndex
    Else
        ParseSectors sx, sy, strSector
        If Len(txtUnit.Text) > 0 And Len(txtQuantity.Text) > 0 Then
            quantity = Val(txtQuantity.Text)
            
            rsBuild.Seek "=", "n", txtUnit.Text
            If Not rsBuild.NoMatch Then
                tweight = rsBuild.Fields("stat3").Value * quantity 'stat3 is lbs field
            End If
            'Compute the path cost
            
            If ParseSectors(nx, ny, txtPath.Text) Then
                pvar = BestPath(strSector, txtPath.Text, MT_COMMODITY)
                bpath = pvar(1)
                pcost = pvar(2)
            Else
                pcost = PathCost(strSector, txtPath.Text, MT_COMMODITY)
                bpath = txtPath.Text
            End If
        End If
    End If
End If
mused = tweight * pcost
lblCost.Caption = "mob. cost: " + CStr(CSng(CLng(CSng(mused) * 1000) / 1000))

rsSectors.Seek "=", sx, sy
If Not rsSectors.NoMatch Then
    mleft = rsSectors.Fields("mob") - mused
    If mleft < 0 Then
        lblLeft.Caption = "Insuff. mobility"
    Else
        lblLeft.Caption = "mob. left: " + CStr(Int(mleft))
    End If
End If

lblBestPath.Caption = "path: " + bpath
lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
If chkDisplayPath.Value = vbChecked Then
    frmDrawMap.DisplayPath SectorString(sx, sy), bpath
End If

rs.Index = hldIndex
End Sub

Private Sub txtPath_GotFocus()
frmDrawMap.DisplayFirstSectorPanel
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And (LCase$(Trim$(Label2.Caption)) = "transport") Then
    KeyAscii = 0
    Load frmPromptWaypoints
    frmPromptWaypoints.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptWaypoints.Width
    frmPromptWaypoints.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptWaypoints.Height) \ 2
    frmPromptWaypoints.Show vbModeless
    frmPromptWaypoints.txtOrigin.SetFocus
End If
End Sub

Private Sub txtUnit_Change()
ComputePathCost
End Sub

Private Sub txtQuantity_Change()
ComputePathCost
End Sub

Private Sub txtUnit_GotFocus()
If Option1 Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
Else
    frmDrawMap.SetUnitDisplay udNUKE, TYPE_ALL, False, False, False
End If
End Sub
