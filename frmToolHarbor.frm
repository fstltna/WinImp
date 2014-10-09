VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmToolHarbor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Harbor at x,y"
   ClientHeight    =   6360
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton cmdBackShort 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Previous Sector that is Short Materials"
      Top             =   5880
      Width           =   375
   End
   Begin VB.OptionButton cmdNextShort 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Next Sector that is Short Materials"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      ToolTipText     =   "Previous Sector"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdPull 
      Caption         =   "Pull Supply"
      Height          =   375
      Left            =   4560
      TabIndex        =   29
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCenter 
      Caption         =   "Center Map"
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   27
      ToolTipText     =   "Next Sector"
      Top             =   5880
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avail Setting"
      Height          =   1125
      Left            =   120
      TabIndex        =   24
      Top             =   5145
      Width           =   1815
      Begin VB.ComboBox cmbUpdates 
         Height          =   315
         ItemData        =   "frmToolHarbor.frx":0000
         Left            =   120
         List            =   "frmToolHarbor.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   740
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Use Update Avail"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   $"frmToolHarbor.frx":0031
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Use Current Avail"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Use the current avail.  Assumes the ships are built then sector's avail is updated."
         Top             =   230
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Builds..."
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource Assumption"
      Height          =   1125
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Worst Case"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Average Case"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   510
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Best Case"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Projected Resources Used"
      Height          =   1335
      Left            =   2040
      TabIndex        =   1
      Top             =   3960
      Width           =   3495
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Remains:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cost: $1,400"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Shortage:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "99 lcms, 99 hcms, 99 guns, 999 avail"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resources Available"
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "hcms"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "lcms"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "guns"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "avail"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9546
      _ExtentY        =   4255
      _Version        =   393216
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Simulated Builds:  1dd, 3bbc, 6awc, 121pt"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2985
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmToolHarbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'280903 rjk: switched consts to enum and add unknown.
'            general reformatting.
'031003 rjk: Added tooltip to cmdBackShort and cmdNextShort.
'171003 rjk: Added Strength fields to Sector database
'241003 rjk: Fixed logic cmdNextShort and cmdBackShort so NextBase is called
'251003 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
'161103 rjk: Changed the name TxtSector to txtSector, code cleanup.
'201103 rjk: Added option to control strength updates
'130304 rjk: Only add a row if there is efficiency gain
'130304 rjk: Fixed FillGridRow to accomodate large avail requirements from St@r W@rs
'180304 rjk: Added Green color if the unit will be at max. eff. gain but not 100%.
'030404 rjk: Switched to use the avail from the server instead of calculating it.
'080404 rjk: Fill in simulated build into Label 4
'080404 rjk: Switched "Left:" to show Not Used: if "Left:" is empty.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'090405 rjk: Added to "Add Update Avail" checkbox to the "Use Current Avail".
'            This can be used to estimate the production for this and the next update.
'200505 rjk: Added capability to forecast upto 3 updates ahead for both avail modes.
'291205 rjk: Added unit workforce to building calculation.
'301205 rjk: Added Theme Game SP 2005, builds the units with military/shells/guns.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'140506 rjk: Use shells for lcms for South Pacific Atlantis
'041006 rjk: Prevent stopped units from appearing in the build tool.
'090607 rjk: Remove the Rollover Avail Scaling by Sector efficiency.
'            Fix Rollover avail calculations for production of avail.

Private ItemType As enumHarborType
Dim secx As Integer
Dim secy As Integer
Dim strBuilds() As String
' Dim cb As Integer    8/2003 efj  removed
Dim OriginChange As Boolean
Dim Costs() As Single
Dim CostItems As Integer
Dim Header As String
Dim Short As Boolean
Dim iMaxLeft As Integer

Private Enum enumHarborType
    FH_UNKNOWN
    FH_BUILD_SHIPS
    FH_BUILD_PLANES
    FH_BUILD_LANDS
    FH_IMPROVE_LANDS
End Enum

Const FH_BEST_CASE = 1
Const FH_WORST_CASE = 2
Const FH_AVERAGE_CASE = 3

Private Sub cmbUpdates_Click()
If Not OriginChange Then
    txtOrigin_Change
End If
End Sub

Private Sub cmdCenter_Click()
frmDrawMap.MoveMap secx, secy
End Sub

Private Sub cmdNext_Click()
NextBase True, False
End Sub

Private Sub cmdBack_Click()
NextBase False, False
End Sub

Private Sub cmdNextShort_Click()
If cmdNextShort.Value Then
    cmdNextShort.Value = False
'Else 102303 rjk: removed so NextBase is called
    NextBase True, True
End If
End Sub

Private Sub cmdBackShort_Click()
If cmdBackShort.Value Then
    cmdBackShort.Value = False
'Else 102303 rjk: removed so NextBase is called
    NextBase False, True
End If
End Sub

Private Sub NextBase(GoingForward As Boolean, ShortOnly As Boolean)
Dim Wrap As Boolean
Dim rs As Recordset

'Go to the current Sector
rsSectors.Seek "=", secx, secy
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If
    
'move to the next sector
If GoingForward Then
    If Not rsSectors.EOF Then
        rsSectors.MoveNext
    End If
Else
    If Not rsSectors.BOF Then
        rsSectors.MovePrevious
    End If
End If

'Search for the next build sector, wraping once if necessary
Wrap = False
While Not Wrap
    If GoingForward And Not Wrap And rsSectors.EOF Then
        rsSectors.MoveFirst
        Wrap = True
    End If
    If Not GoingForward And Not Wrap And rsSectors.BOF Then
        rsSectors.MoveLast
        Wrap = True
    End If
    While Not ((GoingForward And rsSectors.EOF) Or (Not GoingForward And rsSectors.BOF))
        If InStr("h!*f", rsSectors.Fields("des")) > 0 Then
            
            'check for units being built
            Select Case rsSectors.Fields("des")
                Case "h"
                    Set rs = rsShip
                Case "!", "f"
                    Set rs = rsLand
                Case "*"
                    Set rs = rsPlane
            End Select
            
            'Scan units
            If Not (rs.BOF And rs.EOF) Then
                rs.MoveFirst
            End If
            Do While Not rs.EOF
                'Check to see if it is in this hex
                If rsSectors.Fields("x") = rs.Fields("x") And rsSectors.Fields("y") = rs.Fields("y") Then
                    'check to see that it is not 100% already
                    If rs.Fields("eff") < 100 Then
                        txtOrigin.Text = SectorString(rsSectors.Fields("x"), _
                            rsSectors.Fields("y"))
                        If Not (ShortOnly And Not Short) Then
                            Exit Sub
                        Else
                            Exit Do
                        End If
                    End If
                End If
                rs.MoveNext
            Loop
        End If
        If GoingForward Then
            rsSectors.MoveNext
        Else
            rsSectors.MovePrevious
        End If
    Wend
Wend
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdPull_Click()
'Fill resource shorted string
Dim n As Integer
Dim i As Integer
Dim ShortAmt As Integer
Dim ShortItem As String
Dim DistX As Integer
Dim DistY As Integer
Dim strCmd As String

'get sector record
rsSectors.Seek "=", secx, secy
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If
DistX = rsSectors.Fields("dist_x")
DistY = rsSectors.Fields("dist_y")

'get sector record
rsSectors.Seek "=", DistX, DistY
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If

i = 0
For n = 3 To CostItems
    If Costs(6, n) > Costs(5, n) Then
        ShortAmt = Costs(6, n) - Costs(5, n)
        ShortItem = Left$(Trim$(Mid$(Header, (n + 2) * 5 + 1, 5)), 3)
        i = i + 1
        If rsSectors.Fields(ShortItem) >= ShortAmt Then
            strCmd = "move " + ShortItem + " " + SectorString(DistX, DistY) + " " + str$(ShortAmt) + " " + SectorString(secx, secy)
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
    End If
Next n

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

'reset to sector record
rsSectors.Seek "=", secx, secy
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
Dim BuildType As String
Dim i As Integer

Load frmToolHarborQueue
Select Case ItemType
    Case FH_BUILD_SHIPS
        BuildType = "s"
    Case FH_BUILD_LANDS, FH_IMPROVE_LANDS
        BuildType = "l"
    Case FH_BUILD_PLANES
        BuildType = "p"
End Select

'Load Build Queue
On Error GoTo SkipQueue
For i = LBound(strBuilds) + 1 To UBound(strBuilds)
    frmToolHarborQueue.List1.AddItem strBuilds(i)
Next i

frmToolHarborQueue.LoadTypebox BuildType

frmToolHarborQueue.strSector = txtOrigin.Text

SkipQueue: On Error Resume Next
Me.Visible = False
frmToolHarborQueue.Show vbModal

'Simulated Builds:  1dd, 3bbc, 6awc, 121pt
Label4 = vbNullString

'Save Build Queue
ReDim strBuilds(0 To frmToolHarborQueue.List1.ListCount)
For i = 1 To frmToolHarborQueue.List1.ListCount
    Dim n As Integer
    Dim n1 As Integer
    
    strBuilds(i) = frmToolHarborQueue.List1.List(i - 1)
    n = InStr(strBuilds(i), "#")
    If n > 0 Then
        n1 = InStr(n, strBuilds(i), " ")
        If n1 > 0 Then
            Label4 = Label4 + " " + Trim$(Left$(strBuilds(i), n - 1)) + Mid$(strBuilds(i), n, n1 - n)
        End If
    End If
Next i

If Len(Label4) > 0 Then
    Label4 = "Simulated Builds:" + Label4
End If

Unload frmToolHarborQueue
If ItemType <> FH_UNKNOWN Then
    FillGrid
End If
Me.Visible = True
End Sub

Private Sub Form_Load()
'Set always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

lblError.Caption = vbNullString
Label2(0).Caption = vbNullString
Label2(1).Caption = vbNullString
Label2(2).Caption = vbNullString
Label4.Caption = vbNullString

cmbUpdates.ListIndex = 0
If VersionCheck(4, 3, 6) <> VER_LESS Then
    Option2(1).Value = True
    Option2(0).Enabled = False
End If

ReDim strBuilds(0 To 0)
End Sub

Public Sub FillGrid()
Dim Assumptions As Integer
Dim rs As Recordset
Dim BuildType As String
Dim n As Integer
Dim fieldcnt As Integer
Dim MaxEffGain As Integer
Dim strID As String
Dim InitialEff As Integer
Dim ETU_factor As Single
Dim civneeded As Single
Dim iWF As Integer
Dim iOnShip As Integer
Dim iMil As Integer
Dim iIndex As Integer
Dim strField As String

On Error Resume Next

'Reset error string
lblError.Caption = vbNullString

Short = False
ETU_factor = rsVersion.Fields("ETU per update")

'Determmine which type of unit we are displaying
Header = "unit curr new  cost availlcms hcms "
Select Case ItemType
    Case FH_BUILD_SHIPS
        Set rs = rsShip
        rs.Requery
        BuildType = "s"
        fieldcnt = 7
        CostItems = 4
        MaxEffGain = rsVersion.Fields("Eff gain - Ship")
        InitialEff = 20
        Header = Header + "wf   "
        
    Case FH_BUILD_LANDS, FH_IMPROVE_LANDS
        Set rs = rsLand
        rs.Requery
        BuildType = "l"
        fieldcnt = 8
        CostItems = 5
        MaxEffGain = rsVersion.Fields("Eff gain - Land")
        InitialEff = 10
        Header = Header + "guns "
    
    Case FH_BUILD_PLANES
        Set rs = rsPlane
        rs.Requery
        BuildType = "p"
        CostItems = 5
        fieldcnt = 8
        MaxEffGain = rsVersion.Fields("Eff gain - Plane")
        InitialEff = 10
        Header = Header + "mil  "
End Select

ReDim Costs(1 To 7, 1 To CostItems) As Single

'Costs array
'First Dim
'1) holds full cost
'2) holds non rounded cost for current unit
'3) holds rounded (integer) cost for current unit)
'4) holds cumulative remander (used in average case)
'5) holds starting goods
'6) holds the required goods
'7) holds the expended goods

'Second Dim
'1) Cost
'2) Avail
'3) lcm
'4) hcm
'5) mil or guns

'Setup grid
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.Cols = fieldcnt
grid1.ColWidth(0) = -1      'return to default
grid1.ColWidth(-1) = grid1.ColWidth(0) * 0.6
grid1.ColWidth(0) = grid1.ColWidth(0) * 1.5

'Fill in row headers
grid1.col = 0
grid1.row = 0

For n = 0 To fieldcnt - 1
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
    grid1.CellFontBold = True
    grid1.Text = Trim$(Mid$(Header, n * 5 + 1, 5))
Next n

'Clear result labels
Label2(0) = vbNullString
Label2(1) = vbNullString
Label2(2) = vbNullString

'Initial resources
Costs(5, 1) = 9999999        'Assume enough money
Costs(5, 2) = Val(Text1(0))  'avail
Costs(5, 3) = Val(Text1(2))  'lcm
Costs(5, 4) = Val(Text1(3))  'hcm
Costs(5, 5) = Val(Text1(1))  'guns or mil
For n = 1 To CostItems
    Costs(4, n) = 0.5
Next n

'get assumptions
If Option1(0).Value Then
    Assumptions = FH_BEST_CASE
ElseIf Option1(1).Value Then
    Assumptions = FH_AVERAGE_CASE
Else
    Assumptions = FH_WORST_CASE
End If

'Fill in rows with new builds
For n = LBound(strBuilds) + 1 To UBound(strBuilds)
    strID = Left$(strBuilds(n), InStr(strBuilds(n), "(") - 1)
    FillGridRow strID, Costs(), BuildType, Trim$(Left$(strBuilds(n), 5)), _
            Assumptions, 0, InitialEff, 0
Next n

'Fill in rows with actual
rs.MoveFirst
While Not rs.EOF

    'Check to see if it is in this hex
    If secx = rs.Fields("x") And secy = rs.Fields("y") Then
        'check to see that it is not 100% already
        If rs.Fields("eff") < 100 And (VersionCheck(4, 3, 6) = VER_LESS Or rs.Fields("off") = 0) Then
            If BuildType = "p" Then
                iOnShip = rs.Fields("land")
            Else
                iOnShip = -1
            End If
            iWF = WorkForce(BuildType, rs.Fields("type"), rs.Fields("civ"), rs.Fields("mil"), iOnShip) * ETU_factor
            FillGridRow rs.Fields("type") + " #" + CStr(rs.Fields("id")), Costs(), BuildType, rs.Fields("type"), _
                    Assumptions, rs.Fields("eff"), MaxEffGain, iWF
        End If
    End If
    rs.MoveNext
Wend

'Finish new builds
For n = LBound(strBuilds) + 1 To UBound(strBuilds)
    strID = Left$(strBuilds(n), InStr(strBuilds(n), "(") - 1)
    iWF = 0
    If frmOptions.bSP2005 And BuildType = "s" Then
        rsBuild.Seek "=", BuildType, Trim$(Left$(strBuilds(n), 5))
        If Not rsBuild.NoMatch Then
            iMil = 0
            iIndex = 1
            Do
                strField = rsBuild.Fields("cap" + CStr(iIndex))
                If Right$(strField, 1) = "m" Or _
                   (Left$(strField, 1) >= "0" And Left$(strField, 1) <= "9") Then
                    iMil = CInt(Left$(strField, Len(strField) - 1))
                    Exit Do
                End If
                iIndex = iIndex + 1
            Loop While rs.Fields("cap" + CStr(iIndex)) <> "" And iIndex < 10
            iWF = WorkForce(BuildType, Trim$(Left$(strBuilds(n), 5)), 0, iMil, False) * ETU_factor
        End If
    End If
    FillGridRow strID, Costs(), BuildType, Trim$(Left$(strBuilds(n), 5)), _
            Assumptions, InitialEff, MaxEffGain, iWF
Next n

'Fill resource expended string
For n = 2 To CostItems
    If Costs(7, n) > 0 Then
        Label2(0) = Label2(0) + CStr(Costs(7, n)) + " " + Trim$(Mid$(Header, (n + 2) * 5 + 1, 5)) + " "
    End If
Next n

If Len(Label2(0)) = 0 Then   'make sure something was built
    Label2(0) = "Nothing Built"
End If
'Fill cost string
Label2(2).Caption = "Cost: $" + CStr(Costs(7, 1))

'Fill resource shorted string
For n = 2 To CostItems
    If Costs(6, n) > Costs(5, n) Then
        Label2(1) = Label2(1) + CStr(Costs(6, n) - Costs(5, n)) + " " + Trim$(Mid$(Header, (n + 2) * 5 + 1, 5)) + " "
        Short = True
        If n = 2 And HarborAvail Then 'for avail - calc civs needed
            civneeded = (Costs(6, n) - Costs(5, n)) / (ETU_factor / 100)
            civneeded = civneeded / (1 + (ETU_factor * rsVersion.Fields("Civ Birthrate") / 1000))
            Label2(1) = Label2(1) + "(" + CStr(Int(civneeded + 0.9)) + " civs) "
        End If
    End If
Next n
If Len(Label2(1)) = 0 Then
    Label2(1) = "None"
Else
     Label2(2) = Label2(2) + " ($" + CStr(Costs(6, 1)) + " for full build)"
End If
Label2(1) = "Short: " + Label2(1)

'Fill resource left string
Label2(3) = vbNullString
For n = 2 To CostItems
    If Costs(6, n) > Costs(7, n) Then
        Label2(3) = Label2(3) + CStr(Costs(6, n) - Costs(7, n)) + " " + Trim$(Mid$(Header, (n + 2) * 5 + 1, 5)) + " "
    End If
Next n
If Len(Label2(3)) = 0 Then
    For n = 2 To CostItems
        If Costs(5, n) > Costs(7, n) Then
            If n = 2 And iMaxLeft < CStr(Costs(5, n) - Costs(7, n)) Then
                Label2(3) = Label2(3) + CStr(iMaxLeft)
            Else
                Label2(3) = Label2(3) + CStr(Costs(5, n) - Costs(7, n))
            End If
            Label2(3) = Label2(3) + " " + Trim$(Mid$(Header, (n + 2) * 5 + 1, 5)) + " "
        End If
    Next n
    Label2(3) = "Not used: " + Label2(3)
Else
    Label2(3) = "Left: " + Label2(3)
End If

Error_Exit:
'Reshow the grid
grid1.Visible = True
End Sub

Private Sub FillGridRow(UnitDes As String, ByRef Costs() As Single, BuildType As String, UnitType As String, Assumptions As Integer, _
                        ByVal CurrEff As Integer, ByVal MaxEffGain As Integer, iWF As Integer)
Dim OutofResources As Boolean
Dim NewEff As Integer
Dim n As Integer
Dim X As Long '130304 rjk: Changed to Long to Large avail requirements from St@r W@rs
Dim EffGain As Single
Dim hldcosts(1 To 2, 1 To 5) As Single
Dim cRowColor As Long
Dim fWF As Single

fWF = CSng(iWF) / 100#
'get the build requirements
rsBuild.Seek "=", BuildType, UnitType
If rsBuild.NoMatch Then
    lblError.Caption = "Type " + UnitType + " unknown"
Else
    'Compute efficieny gain
    NewEff = CurrEff + MaxEffGain
    If NewEff > 100 Then
        NewEff = 100
    End If
    EffGain = (NewEff - CurrEff) / 100
    
    'load costs
    Costs(1, 1) = rsBuild.Fields("cost")
    '030404 rjk: Switched to use the avail from the server instead of calculating it,
    '            in case the calculation in the server is changed.
    '            20 + rsBuild.Fields("lcm") + rsBuild.Fields("hcm") * 2
    Costs(1, 2) = rsBuild.Fields("avail")
    Costs(1, 3) = rsBuild.Fields("lcm")
    Costs(1, 4) = rsBuild.Fields("hcm")
    Select Case ItemType
    Case FH_BUILD_SHIPS
    Case FH_BUILD_PLANES
        Costs(1, 5) = rsBuild.Fields("mil")
    Case FH_BUILD_LANDS, FH_IMPROVE_LANDS
        Costs(1, 5) = rsBuild.Fields("gun")
    End Select
    
    'compute actual costs
    ComputeCosts Costs, EffGain, Assumptions, 0
    
    'check for overages (not enough resouces)
    OutofResources = False
    For n = 2 To CostItems
        X = Costs(5, n) - Costs(7, n)
        If n = 2 Then
            X = X + fWF
        End If
        If X <= 0 And Costs(3, n) > 0 Then
            OutofResources = True
        End If
    Next n
    
    'if we are not out of anything, check that we have enough
    'for a full build
    If Not OutofResources Then
        For n = 2 To CostItems
            X = Costs(5, n) - Costs(7, n)
            If n = 2 Then
                X = X + fWF
            End If
            If Costs(3, n) > X And Costs(3, n) > 0 Then
                OutofResources = True
                If (X / Costs(1, n)) < EffGain Then
                    EffGain = Int((X * 100) / Costs(1, n)) / 100
                End If
            End If
        Next n
        
        'If this entry caused you to go over
        If OutofResources Then
            'Save full costs that will be spent
            For n = 1 To CostItems
                hldcosts(1, n) = Costs(3, n)
                hldcosts(2, n) = Costs(4, n)
            Next n
            
            'recompute how much can be build without going over
            ComputeCosts Costs, EffGain, Assumptions, fWF
           
            If EffGain > 0 Then

                'Write out a special row
                grid1.row = grid1.row + 1
                
                'write out the description
                grid1.col = 0
                grid1.CellForeColor = vbBlack
                grid1.Text = UnitDes
                grid1.CellAlignment = flexAlignLeftCenter
                
                'old efficiency
                grid1.col = 1
                grid1.Text = CStr(CurrEff)
                
                'new efficency
                grid1.col = 2
                grid1.Text = CStr(CurrEff + CInt(EffGain * 100))
                'write out values
                For n = 1 To CostItems
                'Save actual costs that will be spent
                    Costs(7, n) = Costs(7, n) + Costs(3, n)
                    Costs(6, n) = Costs(6, n) + Costs(3, n)
                    grid1.col = n + 2
                    grid1.Text = CStr(CInt(Costs(3, n)))
                    If grid1.Text = "0" Then
                        grid1.Text = vbNullString
                    End If
                    grid1.CellAlignment = flexAlignRightCenter
                Next n
                grid1.Rows = grid1.Rows + 1 '130304 rjk: Only add a row if there is efficiency gain
            End If
            
            'now reset current eff and costs
            CurrEff = CurrEff + CInt(EffGain * 100)
            EffGain = (NewEff - CurrEff) / 100
            For n = 1 To CostItems
                Costs(3, n) = hldcosts(1, n) - Costs(3, n)
                Costs(4, n) = hldcosts(2, n)
            Next n
        End If
    End If
    grid1.row = grid1.row + 1
    
    If (CurrEff + CInt(EffGain * 100)) >= 100 Then
        cRowColor = vbBlue
    Else
        cRowColor = vbMediumGreen
    End If
    If OutofResources Then
        cRowColor = vbRed
    End If
    
    'write out the description
    grid1.col = 0
    grid1.CellForeColor = cRowColor
    
    grid1.Text = UnitDes ' Set in case lookup fails
    grid1.CellAlignment = flexAlignLeftCenter
    
    'old efficiency
    grid1.col = 1
    grid1.CellForeColor = cRowColor
    grid1.Text = CStr(CurrEff)
    
    'new efficency
    grid1.col = 2
    grid1.CellForeColor = cRowColor
    grid1.Text = CStr(NewEff)
    
    'write out values
    For n = 1 To CostItems
        grid1.col = n + 2
        Costs(6, n) = Costs(6, n) + Costs(3, n)
        If Not OutofResources Then
            Costs(7, n) = Costs(7, n) + Costs(3, n)
        End If
        grid1.CellForeColor = cRowColor
        'negatives are in red
        If Costs(6, n) > Costs(5, n) Then
            grid1.CellForeColor = vbRed
        ElseIf OutofResources Then
            grid1.CellForeColor = RGB(127, 127, 127)
        End If
        grid1.Text = CStr(CInt(Costs(3, n)))
        If grid1.Text = "0" Then
            grid1.Text = vbNullString
        End If
        grid1.CellAlignment = flexAlignRightCenter
    Next n
    grid1.Rows = grid1.Rows + 1
End If
End Sub

Private Sub ComputeCosts(ByRef Costs() As Single, EffGain As Single, Assumption As Integer, fWF As Single)
Dim n As Integer

For n = 1 To CostItems
    If n = 2 And fWF > 0# Then
        If EffGain > (fWF / Costs(1, n)) Then
            Costs(2, n) = Costs(1, n) * (EffGain - (fWF / Costs(1, n)))
        Else
            Costs(2, n) = 0
        End If
    Else
        Costs(2, n) = Costs(1, n) * EffGain
    End If
    If Assumption = FH_BEST_CASE Then
        Costs(3, n) = CSng(Int(Costs(2, n)))
    ElseIf Assumption = FH_WORST_CASE Then
        Costs(3, n) = CSng(Int(Costs(2, n) + 0.995))
    Else
        Costs(3, n) = CSng(Int(Costs(2, n)))
        Costs(4, n) = Costs(4, n) + Costs(2, n) - Costs(3, n)
        If Costs(4, n) > 1 Then
            Costs(3, n) = Costs(3, n) + 1
            Costs(4, n) = Costs(4, n) - 1
        End If
    End If
Next n
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click(Index As Integer)
If ItemType <> FH_UNKNOWN Then
    FillGrid
End If
End Sub

Private Sub Option2_Click(Index As Integer)
If Not OriginChange Then
    'Toggle the avail settings
    Select Case ItemType
    Case FH_BUILD_SHIPS
        HarborAvail = Option2(1)
    Case FH_BUILD_LANDS
        HQAvail = Option2(1)
    Case FH_IMPROVE_LANDS
        FortAvail = Option2(1)
    Case FH_BUILD_PLANES
        AirportAvail = Option2(1)
    End Select
    
    SaveNationVariables
    txtOrigin_Change
End If
End Sub

Private Sub Text1_Change(Index As Integer)
If ItemType <> FH_UNKNOWN Then
    FillGrid
End If
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

'use origin change to avoid firing cascading events
OriginChange = True

ReDim strBuilds(0 To 0)
Label4 = vbNullString

If ParseSectors(sx, sy, txtOrigin.Text) Then
    secx = sx
    secy = sy
    FillSectorInfo
    If ItemType <> FH_UNKNOWN Then
        FillGrid
    End If
Else
    HandleError "Not a valid Sector"
End If
OriginChange = False
End Sub

Private Sub FillSectorInfo()
Dim iCurrentAvail As Integer
Dim iRolloverAvail As Integer
Dim iUpdateAvail As Integer
Dim v As productionDataType

ItemType = FH_UNKNOWN

'get sector record
rsSectors.Seek "=", secx, secy
If rsSectors.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If

'set sector type
Select Case rsSectors.Fields("des")
Case "h"
    ItemType = FH_BUILD_SHIPS
    Me.Caption = "Harbor at " + SectorString(secx, secy)
    Text1(1).Visible = False
    Label1(1).Visible = False
    Command1.Enabled = True
    Option2(0) = Not HarborAvail
    Option2(1) = HarborAvail

Case "!"
    ItemType = FH_BUILD_LANDS
    Me.Caption = "Headquarters at " + SectorString(secx, secy)
    Text1(1).Visible = True
    Label1(1).Visible = True
    Text1(1).Text = CStr(rsSectors.Fields("gun"))
    Label1(1).Caption = "guns"
    Command1.Enabled = True
    Option2(0) = Not HQAvail
    Option2(1) = HQAvail

Case "*"
    ItemType = FH_BUILD_PLANES
    Me.Caption = "Airfield at " + SectorString(secx, secy)
    Text1(1).Visible = True
    Label1(1).Visible = True
    Text1(1).Text = CStr(rsSectors.Fields("mil"))
    Label1(1).Caption = "mil"
    Command1.Enabled = True
    Option2(0) = Not AirportAvail
    Option2(1) = AirportAvail

Case "f"
    ItemType = FH_IMPROVE_LANDS
    Me.Caption = "Fortress at " + SectorString(secx, secy)
    Text1(1).Visible = True
    Label1(1).Visible = True
    Text1(1).Text = CStr(rsSectors.Fields("gun"))
    Label1(1).Caption = "guns"
    Command1.Enabled = False
    Option2(0) = Not FortAvail
    Option2(1) = FortAvail

Case Else
    HandleError "Sector not h,!, f, or *"
    Exit Sub    'error, not a sector type
End Select

'Check if avail used is before or after the update
iCurrentAvail = CStr(rsSectors.Fields("avail"))

If Len(Trim$(rsSectors.Fields("sdes"))) = 0 Then
    iRolloverAvail = rsSectors.Fields("avail")
    If iRolloverAvail > rsVersion.Fields("Rollover Avail") Then
        iRolloverAvail = rsVersion.Fields("Rollover Avail")
    End If
Else
    iRolloverAvail = 0
End If

v = Production(rsSectors)
If v.prodValidFlag Then
    iUpdateAvail = CStr(v.ProdAmount)
Else
    iUpdateAvail = iRolloverAvail
End If

If Option2(0) Then
    If cmbUpdates.ItemData(cmbUpdates.ListIndex) = 1 Then
        iMaxLeft = iCurrentAvail
        Text1(0).Text = iCurrentAvail
    Else
        If v.prodValidFlag Then
            iMaxLeft = iUpdateAvail
            Text1(0).Text = CStr(iCurrentAvail + iRolloverAvail + _
                ((iUpdateAvail - iRolloverAvail) * (cmbUpdates.ItemData(cmbUpdates.ListIndex) - 1)))
        Else
            iMaxLeft = iCurrentAvail
            Text1(0).Text = CStr(iCurrentAvail)
        End If
    End If
Else
    If v.prodValidFlag Then
        iMaxLeft = iUpdateAvail
        Text1(0).Text = CStr(iRolloverAvail + _
            ((iUpdateAvail - iRolloverAvail) * cmbUpdates.ItemData(cmbUpdates.ListIndex)))
    Else
        iMaxLeft = iRolloverAvail
        Text1(0).Text = iRolloverAvail
    End If
End If

If frmOptions.bSPAtlantis Then
    Text1(2).Text = CStr(rsSectors.Fields("shell"))
Else
    Text1(2).Text = CStr(rsSectors.Fields("lcm"))
End If
Text1(3).Text = CStr(rsSectors.Fields("hcm"))
End Sub

Private Sub HandleError(ErrMsg As String)
lblError.Caption = ErrMsg
Text1(0).Text = vbNullString
Text1(1).Text = vbNullString
Text1(2).Text = vbNullString
Text1(3).Text = vbNullString
Label2(0).Caption = vbNullString
Label2(1).Caption = vbNullString
Label2(2).Caption = vbNullString
grid1.Clear
End Sub

Private Function WorkForce(strBuildType As String, strType As String, iCiv As Integer, iMil As Integer, iOnShip As Integer) As Integer
Dim hldIndex As String

WorkForce = 0
hldIndex = ""

On Error GoTo WorkForce_error
Select Case strBuildType
Case "s"
    If Not (rsBuild.BOF And rsBuild.EOF) Then
        rsBuild.Seek "=", strBuildType, strType
        If Not rsBuild.NoMatch Then
            If rsBuild.Fields("stat6") > 0 Then
                WorkForce = iMil / 2
            Else
                WorkForce = iCiv / 2 + iMil / 5
            End If
        End If
    End If
Case "p"
    If iOnShip <> -1 Then
        If Not (rsShip.BOF And rsShip.EOF) Then
            hldIndex = rsShip.Index
            rsShip.Index = "PrimaryKey"
            rsShip.Seek "=", iOnShip
            If Not rsShip.NoMatch Then
                WorkForce = rsShip.Fields("mil") / 2
            End If
        End If
    End If
End Select

WorkForce = WorkForce * cmbUpdates.ItemData(cmbUpdates.ListIndex)
If hldIndex <> "" Then
    rsShip.Index = hldIndex
End If
Exit Function

WorkForce_error:
If hldIndex <> "" Then
    rsShip.Index = hldIndex
End If
WorkForce = 0
End Function
