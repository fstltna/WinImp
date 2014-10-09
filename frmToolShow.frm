VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolShow 
   Caption         =   "Show Tool"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   336
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   593
      ButtonWidth     =   714
      ButtonHeight    =   503
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ship"
            Object.ToolTipText     =   "Show Ships"
            ImageIndex      =   3
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Land"
            Object.ToolTipText     =   "Show Lands"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Plane"
            Object.ToolTipText     =   "Show Planes"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuke"
            Object.ToolTipText     =   "Show Nukes"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sector"
            Object.ToolTipText     =   "Show Sector"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bridge"
            Object.ToolTipText     =   "Show Bridges and Towers"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Item"
            Object.ToolTipText     =   "Show Items"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Build"
            Object.ToolTipText     =   "Show Builds"
            ImageIndex      =   8
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stat"
            Object.ToolTipText     =   "Show Statistics"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cap"
            Object.ToolTipText     =   "Show Capacities"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   495
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTech 
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      Text            =   "999"
      ToolTipText     =   "Optional Tech Level for Refresh"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefreshUnits 
      Caption         =   "Refresh Units"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefreshAll 
      Caption         =   "Refresh All"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox TextBox1 
      Height          =   1215
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Width           =   855
      _ExtentX        =   1503
      _ExtentY        =   2138
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmToolShow.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   7853
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5385
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   445
      Style           =   1
      SimpleText      =   "Printing for tech level"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":0081
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":05D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":0B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":1083
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":2581
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":289B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":2BB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":2ECF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":3161
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolShow.frx":347B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Refresh Tech "
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "frmToolShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'Change Log:
'051203 rjk: removed the special case for SHOW_CAP in determining what tech. level to display
'            now all reports will be at a consisent level.
'121303 rjk: Added show item. Fixed sorting for SHOW_CAPACITY in SHOW_SECTOR
'250104 rjk: Added uw as cargo for St@r W@rs
'270304 rjk: Switched over to DeleteAllRecords for clearing a table
'261105 rjk: Do not compute avail use the value from the server.
'281205 rjk: Switch Infrastructure from text to rsBuild table and present in a table.
'180206 rjk: Add research column for show nukes.
'120306 rjk: Add support for xdump in Refresh actions.
'180606 rjk: Updated sector mobility costs for 4.3.6 server changes.
'140706 rjk: Fixed FillGridItem, it was showing one too many rows.

Dim SortCol As Integer
Dim SortDecending As Boolean
Dim UnitType As String
Dim ReportType As String
Dim ReportTech As Integer

Enum sort_type
    TextSort = 0
    NumericSort = 1
End Enum

Const SHOW_SHIP = "s"
Const SHOW_PLANE = "p"
Const SHOW_LAND = "l"
Const SHOW_NUKE = "n"
Const SHOW_SECTOR = "c"
Const SHOW_BRIDGE = "b"
Const SHOW_TOWER = "t"
Const SHOW_ITEM = "i"

Const SHOW_BUILD = "b"
Const SHOW_STAT = "s"
Const SHOW_CAPACITY = "c"

Private Sub FillTitle()
Dim strTemp As String

strTemp = "Show "
Select Case UnitType
Case SHOW_SHIP
    strTemp = strTemp + "Ship "
Case SHOW_PLANE
    strTemp = strTemp + "Plane "
Case SHOW_LAND
    strTemp = strTemp + "Land "
Case SHOW_NUKE
    strTemp = strTemp + "Nuke "
Case SHOW_SECTOR
    strTemp = strTemp + "Sector "
Case SHOW_BRIDGE
    strTemp = strTemp + "Bridge/Bridge Tower "
Case SHOW_ITEM
    strTemp = strTemp + "Item "
End Select

Select Case ReportType
Case SHOW_BUILD
    strTemp = strTemp + "Build"
Case SHOW_STAT
    strTemp = strTemp + "Statistics"
Case SHOW_CAPACITY
    strTemp = strTemp + "Capacities"
End Select

Me.Caption = strTemp
End Sub

Public Sub FillGrid()
FillTitle
SetReportTech
Select Case ReportType
Case SHOW_BUILD
    Select Case UnitType
        Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE, SHOW_NUKE
            FillGridBuild
        Case SHOW_SECTOR
            FillGridBuildSector
        Case SHOW_BRIDGE
            FillTextBox
        Case SHOW_ITEM
            FillGridItem
    End Select
Case SHOW_STAT
    Select Case UnitType
        Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE, SHOW_NUKE
            FillGridStat
        Case SHOW_BRIDGE
            FillTextBox
        Case SHOW_SECTOR
            FillGridSector
        Case SHOW_ITEM
            FillGridItem
    End Select
Case SHOW_CAPACITY
    Select Case UnitType
        Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE
            FillGridCap
        Case SHOW_NUKE
            FillGridStat
        Case SHOW_BRIDGE
            FillTextBox
        Case SHOW_SECTOR
            FillGridSector
        Case SHOW_ITEM
            FillGridItem
    End Select
End Select
End Sub

Private Sub FillGridBuild()
Dim n As Integer
Dim l As Long
Dim MaxColumn As Integer
Dim AvailCol As Integer
Dim hldIndex As Variant

hldIndex = rsBuild.Index
rsBuild.Index = "tech"

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.Cols = 6
grid1.FixedCols = 1
grid1.FixedRows = 1

Select Case UnitType
Case SHOW_SHIP
    MaxColumn = 6
    AvailCol = 4
Case SHOW_LAND
    MaxColumn = 7
    AvailCol = 5
    grid1.TextMatrix(0, AvailCol - 1) = "guns"
Case SHOW_PLANE
    MaxColumn = 7
    AvailCol = 5
    grid1.TextMatrix(0, AvailCol - 1) = "crew"
Case SHOW_NUKE
    MaxColumn = 9
    AvailCol = 6
    grid1.TextMatrix(0, AvailCol - 2) = "oil"
    grid1.TextMatrix(0, AvailCol - 1) = "rad"
End Select

grid1.Cols = MaxColumn + 1
'Fill headers
grid1.TextMatrix(0, 0) = "type"
grid1.TextMatrix(0, 1) = "description"
grid1.TextMatrix(0, 2) = "lcm"
grid1.TextMatrix(0, 3) = "hcm"
grid1.TextMatrix(0, AvailCol) = "avail"
grid1.TextMatrix(0, AvailCol + 1) = "tech"
If UnitType = SHOW_NUKE Then
    grid1.TextMatrix(0, AvailCol + 2) = "res"
    grid1.TextMatrix(0, AvailCol + 3) = "cost"
Else
    grid1.TextMatrix(0, AvailCol + 2) = "cost"
End If

grid1.ColWidth(-1) = -1
l = grid1.ColWidth(0) / 2
grid1.ColWidth(0) = l
grid1.ColWidth(1) = l * 3
grid1.ColData(0) = TextSort
grid1.ColData(1) = TextSort
grid1.row = 0
For n = 2 To MaxColumn
    grid1.ColWidth(n) = l
    grid1.ColData(n) = NumericSort
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
Next n
grid1.ColWidth(MaxColumn) = l * 1.5
'Fill in rows
n = 1
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = UnitType Then
            grid1.TextMatrix(n, 0) = rsBuild.Fields("id").Value
            grid1.TextMatrix(n, 1) = rsBuild.Fields("desc").Value
            grid1.TextMatrix(n, 2) = rsBuild.Fields("lcm").Value
            grid1.TextMatrix(n, 3) = rsBuild.Fields("hcm").Value
            grid1.TextMatrix(n, AvailCol) = rsBuild.Fields("avail")
            grid1.TextMatrix(n, AvailCol + 1) = rsBuild.Fields("tech").Value
            grid1.TextMatrix(n, AvailCol + 2) = rsBuild.Fields("cost").Value
            If UnitType = SHOW_LAND Then
                grid1.TextMatrix(n, AvailCol - 1) = rsBuild.Fields("gun").Value
            ElseIf UnitType = SHOW_PLANE Then
                grid1.TextMatrix(n, AvailCol - 1) = rsBuild.Fields("mil").Value
            ElseIf UnitType = SHOW_NUKE Then
                grid1.TextMatrix(n, AvailCol - 1) = rsBuild.Fields("rad").Value
                grid1.TextMatrix(n, AvailCol - 2) = rsBuild.Fields("oil").Value
                grid1.TextMatrix(n, AvailCol + 2) = rsBuild.Fields("stat5").Value
                grid1.TextMatrix(n, AvailCol + 3) = rsBuild.Fields("cost").Value
            End If
            
            grid1.row = n
            If rsBuild.Fields("tech") > TechLevel Then
                grid1.row = n
                For l = 0 To MaxColumn
                    grid1.col = l
                    grid1.CellForeColor = vbRed
                Next
            End If
            
            For l = 2 To MaxColumn
                grid1.col = l
                grid1.CellAlignment = flexAlignRightCenter
            Next l
            
            n = n + 1
            grid1.Rows = n + 1
        End If
        rsBuild.MoveNext
    Wend
End If
grid1.Visible = True

rsShowText.Seek "=", UnitType, "b", 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If

rsBuild.Index = hldIndex
End Sub

Private Sub FillGridStat()
Dim n As Integer
Dim l As Long
Dim MaxColumn As Integer
Dim hldIndex As Variant

hldIndex = rsBuild.Index
rsBuild.Index = "tech"

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.FixedCols = 1
grid1.FixedRows = 1

Select Case UnitType
Case SHOW_SHIP
    MaxColumn = 11
    grid1.Cols = MaxColumn + 1
    grid1.TextMatrix(0, 2) = "def."
    grid1.TextMatrix(0, 3) = "spd"
    grid1.TextMatrix(0, 4) = "visib."
    grid1.TextMatrix(0, 5) = "spy"
    grid1.TextMatrix(0, 6) = "range"
    grid1.TextMatrix(0, 7) = "fire"
    grid1.TextMatrix(0, 8) = "land"
    grid1.TextMatrix(0, 9) = "plane"
    grid1.TextMatrix(0, 10) = "helo"
    grid1.TextMatrix(0, 11) = "xplane"
Case SHOW_LAND
    MaxColumn = 17
    grid1.Cols = MaxColumn + 1
    grid1.TextMatrix(0, 2) = "att."
    grid1.TextMatrix(0, 3) = "def."
    grid1.TextMatrix(0, 4) = "vul."
    grid1.TextMatrix(0, 5) = "spd"
    grid1.TextMatrix(0, 6) = "visib."
    grid1.TextMatrix(0, 7) = "spy"
    grid1.TextMatrix(0, 8) = "radius"
    grid1.TextMatrix(0, 9) = "range"
    grid1.TextMatrix(0, 10) = "acc."
    grid1.TextMatrix(0, 11) = "fire"
    grid1.TextMatrix(0, 12) = "ammo"
    grid1.TextMatrix(0, 13) = "anti-air"
    grid1.TextMatrix(0, 14) = "f cap"
    grid1.TextMatrix(0, 15) = "f use"
    grid1.TextMatrix(0, 16) = "xplane"
    grid1.TextMatrix(0, 17) = "land"
Case SHOW_PLANE
    MaxColumn = 8
    grid1.Cols = MaxColumn + 1
    grid1.TextMatrix(0, 2) = "acc."
    grid1.TextMatrix(0, 3) = "load"
    grid1.TextMatrix(0, 4) = "att."
    grid1.TextMatrix(0, 5) = "def."
    grid1.TextMatrix(0, 6) = "range"
    grid1.TextMatrix(0, 7) = "fuel"
    grid1.TextMatrix(0, 8) = "stlth%"
Case SHOW_NUKE
    MaxColumn = 4
    grid1.Cols = MaxColumn + 5
    grid1.TextMatrix(0, 2) = "blast"
    grid1.TextMatrix(0, 3) = "dam."
    grid1.TextMatrix(0, 4) = "lbs"
    grid1.TextMatrix(0, 5) = "tech"
    grid1.TextMatrix(0, 6) = "res"
    grid1.TextMatrix(0, 7) = "cost"
    grid1.TextMatrix(0, 8) = "abilities"
End Select
'Fill headers
grid1.TextMatrix(0, 0) = "type"
grid1.TextMatrix(0, 1) = "description"

grid1.ColWidth(-1) = -1
l = grid1.ColWidth(0) / 2
grid1.ColWidth(0) = l
grid1.ColWidth(1) = l * 3
grid1.ColData(0) = TextSort
grid1.ColData(1) = TextSort
grid1.row = 0
For n = 2 To MaxColumn
    grid1.ColWidth(n) = l * 1.25
    grid1.ColData(n) = NumericSort
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
Next n
If UnitType = SHOW_NUKE Then
    grid1.ColWidth(MaxColumn + 1) = l * 1.25
    grid1.col = MaxColumn + 1
    grid1.CellAlignment = flexAlignRightCenter
    grid1.ColWidth(MaxColumn + 2) = l * 1.25
    grid1.col = MaxColumn + 2
    grid1.CellAlignment = flexAlignRightCenter
    grid1.col = MaxColumn + 3
    grid1.CellAlignment = flexAlignRightCenter
    grid1.col = MaxColumn + 4
    grid1.CellAlignment = flexAlignLeftCenter
    grid1.ColData(MaxColumn + 4) = TextSort
End If

'Fill in rows
n = 1
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = UnitType Then
            grid1.TextMatrix(n, 0) = rsBuild.Fields("id").Value
            grid1.TextMatrix(n, 1) = rsBuild.Fields("desc").Value
            For l = 2 To MaxColumn
                If l < 4 And UnitType = SHOW_LAND Then
                    grid1.TextMatrix(n, l) = Format$(rsBuild.Fields("stat" + CStr(l - 1)).Value, "###0.00")
                Else
                    grid1.TextMatrix(n, l) = CStr(rsBuild.Fields("stat" + CStr(l - 1)).Value)
                End If
            Next l
            
            If UnitType = SHOW_NUKE Then
                grid1.TextMatrix(n, MaxColumn + 1) = rsBuild.Fields("tech").Value
                grid1.col = MaxColumn + 1
                grid1.CellAlignment = flexAlignRightCenter
                grid1.TextMatrix(n, MaxColumn + 2) = rsBuild.Fields("stat5").Value
                grid1.col = MaxColumn + 2
                grid1.CellAlignment = flexAlignRightCenter
                grid1.TextMatrix(n, MaxColumn + 3) = rsBuild.Fields("cost").Value
                grid1.col = MaxColumn + 3
                grid1.CellAlignment = flexAlignRightCenter
                grid1.TextMatrix(n, MaxColumn + 4) = rsBuild.Fields("cap1").Value
                grid1.col = MaxColumn + 4
                grid1.CellAlignment = flexAlignLeftCenter
            End If
            
            grid1.row = n
            If rsBuild.Fields("tech") > TechLevel Then
                For l = 0 To MaxColumn
                    grid1.col = l
                    grid1.CellForeColor = vbRed
                Next
                If UnitType = SHOW_NUKE Then
                    For l = MaxColumn + 1 To MaxColumn + 4
                        grid1.col = l
                        grid1.CellForeColor = vbRed
                    Next
                End If
            End If
            
            For l = 2 To MaxColumn
                grid1.col = l
                grid1.CellAlignment = flexAlignRightCenter
            Next l
            
            n = n + 1
            grid1.Rows = n + 1
        End If
        rsBuild.MoveNext
    Wend
End If
grid1.Visible = True

rsShowText.Seek "=", UnitType, ReportType, 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If

rsBuild.Index = hldIndex
End Sub

Private Sub FillGridCap()
Dim n As Integer
Dim l As Long
Dim MaxColumn As Integer
' Dim AvailCol As Integer    8/2003 efj  removed
Dim strTemp As Variant 'drk@unxsoft.com 8/11/02.  Seems to be required to get stuff to co-operate below...
Dim strType As String
Dim CargoColumn As Integer
Dim hldIndex As Variant

hldIndex = rsBuild.Index
rsBuild.Index = "tech"

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.FixedCols = 1
grid1.FixedRows = 1

Select Case UnitType
Case SHOW_SHIP
    MaxColumn = 9
    grid1.Cols = MaxColumn + 1
    CargoColumn = MaxColumn - 1
    grid1.TextMatrix(0, 2) = "mil"
    grid1.TextMatrix(0, 3) = "gun"
    grid1.TextMatrix(0, 4) = "shell"
    grid1.TextMatrix(0, 5) = "food"
    grid1.TextMatrix(0, 6) = "civ"
    grid1.TextMatrix(0, 7) = "uw"
    grid1.TextMatrix(0, 8) = "cargo"
    grid1.TextMatrix(0, 9) = "abilities"
Case SHOW_LAND
    MaxColumn = 6
    grid1.Cols = MaxColumn + 1
    CargoColumn = MaxColumn
    grid1.TextMatrix(0, 2) = "mil"
    grid1.TextMatrix(0, 3) = "gun"
    grid1.TextMatrix(0, 4) = "shell"
    grid1.TextMatrix(0, 5) = "food"
    grid1.TextMatrix(0, 6) = "cargo/abilities"
Case SHOW_PLANE
    MaxColumn = 2
    CargoColumn = 2
    grid1.Cols = MaxColumn + 1
    grid1.TextMatrix(0, 2) = "abilities"
End Select
'Fill headers
grid1.TextMatrix(0, 0) = "type"
grid1.TextMatrix(0, 1) = "description"

grid1.ColWidth(-1) = -1
l = grid1.ColWidth(0) / 2
grid1.ColWidth(0) = l
grid1.ColWidth(1) = l * 3
grid1.ColData(0) = TextSort
grid1.ColData(1) = TextSort
grid1.row = 0
For n = 2 To MaxColumn
    grid1.ColWidth(n) = l
    If n < CargoColumn Then
        grid1.ColData(n) = NumericSort
    Else
        grid1.ColData(n) = TextSort
    End If
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
Next n
grid1.ColWidth(MaxColumn) = l * 10
grid1.col = MaxColumn
grid1.CellAlignment = flexAlignLeftCenter

If UnitType = SHOW_SHIP Then
    grid1.ColWidth(CargoColumn) = l * 3
End If
grid1.col = CargoColumn
grid1.CellAlignment = flexAlignLeftCenter
'Fill in rows
n = 1
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = UnitType Then
            grid1.TextMatrix(n, 0) = rsBuild.Fields("id").Value
            grid1.TextMatrix(n, 1) = rsBuild.Fields("desc").Value
            For l = 2 To CargoColumn - 1
                grid1.TextMatrix(n, l) = "0"
            Next l
            For l = 1 To 20
                strTemp = rsBuild.Fields("cap" + CStr(l)).Value
                If Len(strTemp) > 0 Then
                    If Left$(strTemp, 1) >= "1" And Left$(strTemp, 1) <= "9" Then
                        strType = Right$(strTemp, 1)
                        Select Case strType
                        Case "m"
                            grid1.TextMatrix(n, 2) = Left$(strTemp, Len(strTemp) - 1)
                        Case "g"
                            grid1.TextMatrix(n, 3) = Left$(strTemp, Len(strTemp) - 1)
                        Case "s"
                            grid1.TextMatrix(n, 4) = Left$(strTemp, Len(strTemp) - 1)
                        Case "f"
                            grid1.TextMatrix(n, 5) = Left$(strTemp, Len(strTemp) - 1)
                        Case "c"
                            If UnitType = SHOW_LAND Then '090903 rjk: LOTRII Civilians are allowed as cargo
                                grid1.TextMatrix(n, CargoColumn) = grid1.TextMatrix(n, CargoColumn) + strTemp + " "
                            Else
                                grid1.TextMatrix(n, 6) = Left$(strTemp, Len(strTemp) - 1)
                            End If
                        Case "u"
                            If UnitType = SHOW_LAND Then '240104 rjk: St@r W@rs UW are allowed as cargo
                                grid1.TextMatrix(n, CargoColumn) = grid1.TextMatrix(n, CargoColumn) + strTemp + " "
                            Else
                                grid1.TextMatrix(n, 7) = Left$(strTemp, Len(strTemp) - 1)
                            End If
                        Case Else
                            grid1.TextMatrix(n, CargoColumn) = grid1.TextMatrix(n, CargoColumn) + strTemp + " "
                        End Select
                    Else
                        grid1.TextMatrix(n, MaxColumn) = grid1.TextMatrix(n, MaxColumn) + strTemp + " "
                    End If
                End If
            Next l
            
            grid1.row = n
            If rsBuild.Fields("tech") > TechLevel Then
                For l = 0 To MaxColumn
                    grid1.col = l
                    grid1.CellForeColor = vbRed
                Next
            End If
            
            For l = 2 To MaxColumn - 1
                grid1.col = l
                grid1.CellAlignment = flexAlignRightCenter
            Next l
            grid1.col = CargoColumn
            grid1.CellAlignment = flexAlignLeftCenter
            grid1.col = MaxColumn
            grid1.CellAlignment = flexAlignLeftCenter
            
            n = n + 1
            grid1.Rows = n + 1
        End If
        rsBuild.MoveNext
    Wend
End If
grid1.Visible = True

rsShowText.Seek "=", UnitType, "c", 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If

rsBuild.Index = hldIndex
End Sub

Private Sub cmdClear_Click()
Dim Result As Integer

Result = MsgBox("Do you want to clear ALL unit data " + _
        "stored for this game, including higher tech defaults?", vbYesNo, "Clear Confirmation")

If Result = vbYes Then
    DeleteAllRecords rsBuild

    If Not (rsShowText.EOF And rsShowText.BOF) Then
        rsShowText.MoveFirst
    End If
    
    While Not rsShowText.EOF
        If InStr("spln", rsShowText.Fields("unittype")) > 0 Then
            If rsShowText.Fields("linenumber") = 1 Then
                rsShowText.Edit
                rsShowText.Fields("text") = "No data available - use Refresh"
                rsShowText.Update
            Else
                rsShowText.Delete
            End If
        End If
        rsShowText.MoveNext
    Wend
End If

Result = MsgBox("Do you want to clear sector and bridge data " + _
        "stored for this game.  If you choose Yes, you will need to Refresh sector " + _
        " data for all of WinACE's functions to work correctly", vbYesNo, "Clear Confirmation")

If Result = vbYes Then
    DeleteAllRecords rsSectorType

    If Not (rsShowText.EOF And rsShowText.BOF) Then
        rsShowText.MoveFirst
    End If
    
    While Not rsShowText.EOF
        If InStr("cbt", rsShowText.Fields("unittype")) > 0 Then
            If rsShowText.Fields("linenumber") = 1 Then
                rsShowText.Edit
                rsShowText.Fields("text") = "No data available - use Refresh"
                rsShowText.Update
            Else
                rsShowText.Delete
            End If
        End If
        rsShowText.MoveNext
    Wend
End If
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Function BuildCommand(unit As String, report As String) As String
Select Case unit
Case SHOW_LAND
    BuildCommand = "land"
Case SHOW_PLANE
    BuildCommand = "plane"
Case SHOW_SHIP
    BuildCommand = "ship"
Case SHOW_BRIDGE
    BuildCommand = "bridge"
Case SHOW_TOWER
    BuildCommand = "tower"
Case SHOW_NUKE
    BuildCommand = "nuke"
Case SHOW_SECTOR
    BuildCommand = "sector"
Case SHOW_ITEM
    BuildCommand = "item"
End Select
BuildCommand = "show " + BuildCommand + " " + report + " "
If Len(txtTech.Text) > 0 Then
    BuildCommand = BuildCommand + " " + txtTech.Text
' ElseIf report = SHOW_STAT Then 'rjk 05/12/03: removed this special case, so all reports are displayed at a consisent tech. level.
'    BuildCommand = BuildCommand + " " + CStr(TechLevel)
End If
End Function

Private Function BuildXdumpCommand(unit As String) As String
Select Case unit
Case SHOW_LAND
    BuildXdumpCommand = "land-chr"
Case SHOW_PLANE
    BuildXdumpCommand = "plane-chr"
Case SHOW_SHIP
    BuildXdumpCommand = "ship-chr"
Case SHOW_BRIDGE
    BuildXdumpCommand = "bridge-chr"
Case SHOW_TOWER
    BuildXdumpCommand = "tower"
Case SHOW_NUKE
    BuildXdumpCommand = "nuke-chr"
Case SHOW_SECTOR
    BuildXdumpCommand = "sector-chr"
Case SHOW_BRIDGE, SHOW_TOWER
    BuildXdumpCommand = "ver"
Case SHOW_ITEM
    BuildXdumpCommand = "item"
End Select
BuildXdumpCommand = "xdump " + BuildXdumpCommand + " *"
End Function

Private Sub cmdRefresh_Click()
If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand BuildCommand(UnitType, ReportType), True
    If UnitType = SHOW_BRIDGE Then
        frmEmpCmd.SubmitEmpireCommand BuildCommand(SHOW_TOWER, ReportType), True
    End If
Else
    frmEmpCmd.SubmitEmpireCommand BuildXdumpCommand(UnitType), True
End If
End Sub

Public Sub cmdRefreshAll_Click()
If VersionCheck(4, 3, 0) = VER_LESS Then
    cmdRefreshUnits_Click
    frmEmpCmd.SubmitEmpireCommand BuildCommand("n", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("n", "s"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("n", "c"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("c", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("c", "s"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("c", "c"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("b", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("t", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("i", "b"), False
Else
    RequestConfigurationTables
End If
End Sub

Private Sub cmdRefreshUnits_Click()
If VersionCheck(4, 3, 0) = VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand BuildCommand("s", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("s", "s"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("s", "c"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("l", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("l", "s"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("l", "c"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("p", "b"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("p", "s"), False
    frmEmpCmd.SubmitEmpireCommand BuildCommand("p", "c"), False
Else
    RequestUnitConfigurationTables
End If
End Sub

Private Sub Form_Load()

'Set always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

'Restore Toolbar
Toolbar1.RestoreToolbar APPNAME, "Layout", "ToolShow Toolbar"

Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2


UnitType = SHOW_SHIP
ReportType = SHOW_BUILD
txtTech.Text = vbNullString
FillGrid

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then Exit Sub

If Me.Width < cmdOK.Width * 5 Then
    Me.Width = cmdOK.Width * 5
End If

If Me.Height < cmdOK.top + cmdOK.Height * 3 Then
    Me.Height = cmdOK.top + cmdOK.Height * 3
End If

With Me
    grid1.Move 0, Toolbar1.top + Toolbar1.Height, Width - cmdOK.Width * 1.5, .sb1.top - Toolbar1.Height
    grid1.Refresh
    Toolbar1.Width = .Width
    cmdOK.Left = .Width - cmdOK.Width * 1.25
    cmdRefresh.Left = cmdOK.Left
    cmdRefreshUnits.Left = cmdOK.Left
    cmdRefreshAll.Left = cmdOK.Left
    cmdClear.Left = cmdOK.Left
    Label1.Left = cmdOK.Left
    txtTech.Left = cmdOK.Left + cmdOK.Width - txtTech.Width
End With

With grid1
    TextBox1.Move .Left, .top, .Width, .Height
End With
TextBox1.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
Toolbar1.SaveToolbar APPNAME, "Layout", "ToolShow Toolbar"
End Sub

Private Sub grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If clicking in top row, resort current box
If grid1.MouseRow = 0 Then
    SortGrid grid1.MouseCol
    Exit Sub
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
Select Case Button.key
    Case "Ship"
        UnitType = SHOW_SHIP
    Case "Plane"
        UnitType = SHOW_PLANE
    Case "Land"
        UnitType = SHOW_LAND
    Case "Nuke"
        UnitType = SHOW_NUKE
    Case "Bridge"
        UnitType = SHOW_BRIDGE
    Case "Sector"
        UnitType = SHOW_SECTOR
    Case "Item"
        UnitType = SHOW_ITEM
    Case "Build"
        ReportType = SHOW_BUILD
    Case "Stat"
        ReportType = SHOW_STAT
    Case "Cap"
        ReportType = SHOW_CAPACITY
End Select
        
FillGrid
End Sub

Private Sub FillTextBox()

' Dim n As Integer    8/2003 efj  removed
' Dim l As Long    8/2003 efj  removed
' Dim MaxColumn As Integer    8/2003 efj  removed
' Dim AvailCol As Integer    8/2003 efj  removed

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
TextBox1.Text = vbNullString

If Not (rsShowText.BOF And rsShowText.EOF) Then
    rsShowText.MoveFirst
End If

While Not rsShowText.EOF
    If (rsShowText.Fields("UnitType") = UnitType Or _
        UnitType = SHOW_BRIDGE And rsShowText.Fields("UnitType") = "t") And _
        rsShowText.Fields("reporttype") = SHOW_BUILD And _
        rsShowText.Fields("Linenumber") > 1 Then
        TextBox1.Text = TextBox1.Text + rsShowText.Fields("Text") + vbCrLf
    End If
    rsShowText.MoveNext
Wend

TextBox1.Visible = True
End Sub

Private Sub FillGridBuildSector()
Dim n As Integer
Dim l As Long
Dim MaxColumn As Integer
Dim AvailCol As Integer
Dim hldIndex As Variant

hldIndex = rsBuild.Index
rsBuild.Index = "tech"

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 1
grid1.Cols = 6
grid1.FixedCols = 1
'grid1.FixedRows = 1

'sector type    cost to des    cost for 1% eff   lcms for 1%    hcms for 1%
'Fill First header
grid1.TextMatrix(0, 0) = "sector type"
grid1.TextMatrix(0, 1) = "$$$ - des"
grid1.TextMatrix(0, 2) = "$$$ - 1%"
grid1.TextMatrix(0, 3) = "lcms - 1%"
grid1.TextMatrix(0, 4) = "hcms - 1%"
grid1.TextMatrix(0, 5) = "mobility - 1%"

grid1.ColWidth(-1) = -1
grid1.ColData(0) = TextSort
grid1.col = 0
grid1.row = 0
grid1.CellAlignment = flexAlignLeftCenter
For n = 1 To grid1.Cols - 1
    grid1.ColData(n) = NumericSort
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
Next n
'Fill in rows
n = 1
If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = "i" Then
            grid1.Rows = grid1.Rows + 1
            grid1.TextMatrix(n, 0) = rsBuild.Fields("id").Value
            If Len(grid1.TextMatrix(n, 0)) = 1 Then
                grid1.TextMatrix(n, 1) = rsBuild.Fields("stat1").Value
            Else
                grid1.TextMatrix(n, 1) = "0"
            End If
            grid1.TextMatrix(n, 2) = rsBuild.Fields("cost").Value
            grid1.TextMatrix(n, 3) = rsBuild.Fields("lcm").Value
            grid1.TextMatrix(n, 4) = rsBuild.Fields("hcm").Value
            If Len(grid1.TextMatrix(n, 0)) > 1 Then
                grid1.TextMatrix(n, 5) = rsBuild.Fields("stat2").Value
            Else
                grid1.TextMatrix(n, 5) = "0"
            End If
            n = n + 1
        End If
        rsBuild.MoveNext
    Wend
End If
grid1.Visible = True

rsShowText.Seek "=", UnitType, "b", 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If

rsBuild.Index = hldIndex
End Sub

Private Sub FillGridSector()
Dim n As Integer
Dim l As Long
Dim MaxColumn As Integer

'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.FixedCols = 1
grid1.FixedRows = 1

Select Case ReportType
Case SHOW_CAPACITY
    MaxColumn = 12
    grid1.Cols = MaxColumn + 1
    grid1.TextMatrix(0, 2) = "product"
    grid1.TextMatrix(0, 3) = "use1"
    grid1.TextMatrix(0, 4) = "use2"
    grid1.TextMatrix(0, 5) = "use3"
    grid1.TextMatrix(0, 6) = "level"
    grid1.TextMatrix(0, 7) = "min"
    grid1.TextMatrix(0, 8) = "lag"
    grid1.TextMatrix(0, 9) = "eff%"
    grid1.TextMatrix(0, 10) = "$$$"
    grid1.TextMatrix(0, 11) = "dep"
    grid1.TextMatrix(0, 12) = "code"
Case SHOW_STAT
    MaxColumn = 10
    grid1.Cols = MaxColumn + 1
    If VersionCheck(4, 3, 6) <> VER_LESS Then
        MaxColumn = 11
        grid1.Cols = MaxColumn + 1
        grid1.TextMatrix(0, 2) = "mcost 0%"
        grid1.TextMatrix(0, 3) = "mcost 100%"
        grid1.TextMatrix(0, 4) = "max off"
        grid1.TextMatrix(0, 5) = "max def"
        grid1.TextMatrix(0, 6) = "mil"
        grid1.TextMatrix(0, 7) = "uw"
        grid1.TextMatrix(0, 8) = "civ"
        grid1.TextMatrix(0, 9) = "bar"
        grid1.TextMatrix(0, 10) = "other"
        grid1.TextMatrix(0, 11) = "maxpop"
    Else
        MaxColumn = 10
        grid1.Cols = MaxColumn + 1
        grid1.TextMatrix(0, 2) = "mcost"
        grid1.TextMatrix(0, 3) = "max off"
        grid1.TextMatrix(0, 4) = "max def"
        grid1.TextMatrix(0, 5) = "mil"
        grid1.TextMatrix(0, 6) = "uw"
        grid1.TextMatrix(0, 7) = "civ"
        grid1.TextMatrix(0, 8) = "bar"
        grid1.TextMatrix(0, 9) = "other"
        grid1.TextMatrix(0, 10) = "maxpop"
    End If
End Select

'Fill headers
grid1.TextMatrix(0, 0) = "des"
grid1.TextMatrix(0, 1) = "description"
grid1.row = 0
grid1.col = 0
grid1.CellAlignment = flexAlignCenterCenter
grid1.col = 1
grid1.CellAlignment = flexAlignLeftCenter

grid1.ColWidth(-1) = -1
l = grid1.ColWidth(0) / 2
grid1.ColWidth(0) = l
grid1.ColWidth(1) = l * 3
grid1.ColData(0) = TextSort
grid1.ColData(1) = TextSort
For n = 2 To MaxColumn
    grid1.ColWidth(n) = l * 1.1
    '121303 rjk: Fixed sorting for SHOW_CAPACITY
    If ReportType = SHOW_STAT Or ((n < 7 Or n = 12) And ReportType = SHOW_CAPACITY) Then
        grid1.ColData(n) = TextSort
    Else
        grid1.ColData(n) = NumericSort
    End If
    grid1.col = n
    grid1.CellAlignment = flexAlignRightCenter
Next n

Select Case ReportType
Case SHOW_STAT
    If VersionCheck(4, 3, 6) <> VER_LESS Then
        grid1.ColWidth(2) = l * 1.6
        grid1.ColWidth(3) = l * 2#
        grid1.ColWidth(4) = l * 1.3
        grid1.ColWidth(5) = l * 1.4
        grid1.ColWidth(11) = l * 1.3
    Else
        grid1.ColWidth(3) = l * 1.5
        grid1.ColWidth(4) = l * 1.5
        grid1.ColWidth(10) = l * 1.3
    End If
Case SHOW_CAPACITY
    grid1.ColWidth(2) = l * 1.5
    grid1.col = 2
    grid1.CellAlignment = flexAlignLeftCenter
    grid1.col = 6
    grid1.CellAlignment = flexAlignLeftCenter
    grid1.col = 12
    grid1.CellAlignment = flexAlignCenterCenter
    For n = 7 To 11
        grid1.ColData(n) = NumericSort
    Next n
End Select
'Fill in rows
n = 1
If Not (rsSectorType.BOF And rsSectorType.EOF) Then
    rsSectorType.MoveFirst
End If
    
While Not rsSectorType.EOF
    grid1.TextMatrix(n, 0) = rsSectorType.Fields("des").Value
    grid1.TextMatrix(n, 1) = rsSectorType.Fields("desc").Value
    grid1.row = n
    grid1.col = 0
    grid1.CellAlignment = flexAlignCenterCenter
    grid1.CellFontBold = True
    
    For l = 2 To MaxColumn
        grid1.col = l
        grid1.CellAlignment = flexAlignRightCenter
    Next l
    Select Case ReportType
    Case SHOW_CAPACITY
        grid1.TextMatrix(n, 2) = rsSectorType.Fields("product").Value
        If rsSectorType.Fields("use1n").Value > 0 Then
            grid1.TextMatrix(n, 3) = CStr(rsSectorType.Fields("use1n").Value) + " " + rsSectorType.Fields("use1s").Value
        End If
        If rsSectorType.Fields("use2n").Value > 0 Then
            grid1.TextMatrix(n, 4) = CStr(rsSectorType.Fields("use2n").Value) + " " + rsSectorType.Fields("use2s").Value
        End If
        If rsSectorType.Fields("use3n").Value > 0 Then
            grid1.TextMatrix(n, 5) = CStr(rsSectorType.Fields("use3n").Value) + " " + rsSectorType.Fields("use3s").Value
        End If
        
        grid1.TextMatrix(n, 6) = rsSectorType.Fields("level").Value
        grid1.TextMatrix(n, 7) = rsSectorType.Fields("min").Value
        grid1.TextMatrix(n, 8) = rsSectorType.Fields("lag").Value
        grid1.TextMatrix(n, 9) = rsSectorType.Fields("eff").Value
        grid1.TextMatrix(n, 10) = rsSectorType.Fields("cost").Value
        grid1.TextMatrix(n, 11) = rsSectorType.Fields("dep").Value
        grid1.TextMatrix(n, 12) = rsSectorType.Fields("pcode").Value
        grid1.col = 2
        grid1.CellAlignment = flexAlignLeftCenter
        grid1.col = 3
        grid1.CellAlignment = flexAlignCenterCenter
        grid1.col = 4
        grid1.CellAlignment = flexAlignCenterCenter
        grid1.col = 5
        grid1.CellAlignment = flexAlignCenterCenter
        grid1.col = 6
        grid1.CellAlignment = flexAlignLeftCenter
        grid1.col = 12
        grid1.CellAlignment = flexAlignCenterCenter
    Case SHOW_STAT
        If VersionCheck(4, 3, 6) <> VER_LESS Then
            grid1.TextMatrix(n, 2) = Format$(rsSectorType.Fields("mcost").Value, "##0.000")
            grid1.TextMatrix(n, 3) = Format$(rsSectorType.Fields("mcost100").Value, "##0.000")
            grid1.TextMatrix(n, 4) = Format$(rsSectorType.Fields("offense").Value, "##0.00")
            grid1.TextMatrix(n, 5) = Format$(rsSectorType.Fields("defense").Value, "##0.00")
            grid1.TextMatrix(n, 6) = rsSectorType.Fields("pack_mil").Value
            grid1.TextMatrix(n, 7) = rsSectorType.Fields("pack_uw").Value
            grid1.TextMatrix(n, 8) = rsSectorType.Fields("pack_civ").Value
            grid1.TextMatrix(n, 9) = rsSectorType.Fields("pack_bar").Value
            grid1.TextMatrix(n, 10) = rsSectorType.Fields("pack_other").Value
            grid1.TextMatrix(n, 11) = rsSectorType.Fields("maxpop").Value
        Else
            grid1.TextMatrix(n, 2) = rsSectorType.Fields("mcost").Value
            grid1.TextMatrix(n, 3) = Format$(rsSectorType.Fields("offense").Value, "##0.00")
            grid1.TextMatrix(n, 4) = Format$(rsSectorType.Fields("defense").Value, "##0.00")
            grid1.TextMatrix(n, 5) = rsSectorType.Fields("pack_mil").Value
            grid1.TextMatrix(n, 6) = rsSectorType.Fields("pack_uw").Value
            grid1.TextMatrix(n, 7) = rsSectorType.Fields("pack_civ").Value
            grid1.TextMatrix(n, 8) = rsSectorType.Fields("pack_bar").Value
            grid1.TextMatrix(n, 9) = rsSectorType.Fields("pack_other").Value
            grid1.TextMatrix(n, 10) = rsSectorType.Fields("maxpop").Value
        End If
    End Select
 
    n = n + 1
    grid1.Rows = n + 1
    rsSectorType.MoveNext
Wend
grid1.Visible = True

rsShowText.Seek "=", "c", ReportType, 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If
End Sub

Private Sub SetReportTech()
Dim n As Integer
Dim n2 As Integer

rsShowText.Seek "=", UnitType, ReportType, 1
If rsShowText.NoMatch Then
    ReportTech = TechLevel
    Exit Sub
End If

n = InStr(rsShowText.Fields("Text").Value, "'")
If n = 0 Then
    ReportTech = TechLevel
    Exit Sub
End If

n2 = InStr(n + 1, rsShowText.Fields("Text").Value, "'")

If n2 <= n Then
    ReportTech = TechLevel
    Exit Sub
End If

ReportTech = CInt(Mid$(rsShowText.Fields("Text").Value, n + 1, n2 - n - 1))
End Sub

Private Sub SortGrid(col As Integer)
'This does a simple bubble sort of the grid in place.  Due to the small size
'of the sample, a more efficent sort is usually not necessary.
Dim n As Integer
Dim row As Integer
Dim var1 As Variant
Dim var2 As Variant
Dim sng1 As Single
Dim sng2 As Single

'If clicked multiple times, change the sort order
If SortCol = col Then
    SortDecending = Not SortDecending
End If
SortCol = col

grid1.Visible = False

n = grid1.Rows - 3

If grid1.ColData(col) = TextSort Then
    'text sort
    While n > 0
        For row = 1 To n
            var1 = grid1.TextMatrix(row, col)
            var2 = grid1.TextMatrix(row + 1, col)
            
            If SortDecending Then
                If var1 > var2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            Else
                If var1 < var2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            End If
        Next row
        n = n - 1
    Wend
Else
    'numeric sort
    While n > 0
        For row = 1 To n
            sng1 = CSng(grid1.TextMatrix(row, col))
            sng2 = CSng(grid1.TextMatrix(row + 1, col))
            
            If SortDecending Then
                If sng1 > sng2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            Else
                If sng1 < sng2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            End If
        Next row
        n = n - 1
    Wend
End If
grid1.Visible = True
End Sub

Private Sub FillGridItem()
Dim n As Integer
Dim l As Long
'Reset sort variables
SortCol = -1
SortDecending = True

TextBox1.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = Items.Count + 1
grid1.FixedCols = 1
grid1.FixedRows = 1

grid1.Cols = 9
grid1.row = 0
grid1.ColWidth(-1) = -1
l = grid1.ColWidth(0) / 2
grid1.col = 0
grid1.ColWidth(-1) = l
grid1.TextMatrix(0, 0) = "letter"
grid1.CellAlignment = flexAlignCenterCenter
grid1.ColData(0) = TextSort
grid1.col = 1
grid1.TextMatrix(0, 1) = "value"
grid1.CellAlignment = flexAlignRightCenter
grid1.ColData(1) = NumericSort
grid1.col = 2
grid1.TextMatrix(0, 2) = "sell"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(2) = TextSort
grid1.col = 3
grid1.TextMatrix(0, 3) = "lbs"
grid1.CellAlignment = flexAlignRightCenter
grid1.ColData(3) = NumericSort
grid1.col = 4
grid1.TextMatrix(0, 4) = "rg pk"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(4) = NumericSort
grid1.ColWidth(4) = l * 1.1
grid1.col = 5
grid1.TextMatrix(0, 5) = "wh pk"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(5) = NumericSort
grid1.ColWidth(5) = l * 1.1
grid1.col = 6
grid1.TextMatrix(0, 6) = "ur pk"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(6) = NumericSort
grid1.ColWidth(6) = l * 1.1
grid1.col = 7
grid1.TextMatrix(0, 7) = "bk pk"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(7) = NumericSort
grid1.ColWidth(7) = l * 1.1
grid1.col = 8
grid1.ColWidth(8) = l * 5
grid1.TextMatrix(0, 8) = "name"
grid1.CellAlignment = flexAlignLeftCenter
grid1.ColData(8) = TextSort
'Fill in rows
For n = 1 To Items.Count
    grid1.TextMatrix(n, 0) = Items(n).Letter
    grid1.TextMatrix(n, 1) = CStr(Items(n).Value)
    If Items(n).Sellable Then
        grid1.TextMatrix(n, 2) = "yes"
    Else
        grid1.TextMatrix(n, 2) = "no"
    End If
    grid1.TextMatrix(n, 3) = CStr(Items(n).Weight)
    grid1.TextMatrix(n, 4) = CStr(Items(n).Packing(PACKING_REGULAR))
    grid1.TextMatrix(n, 5) = CStr(Items(n).Packing(PACKING_WAREHOUSE))
    grid1.TextMatrix(n, 6) = CStr(Items(n).Packing(PACKING_URBAN))
    grid1.TextMatrix(n, 7) = CStr(Items(n).Packing(PACKING_BANK))
    grid1.TextMatrix(n, 8) = Items(n).ItemName
Next n

grid1.Visible = True
rsShowText.Seek "=", "c", ReportType, 1
If Not rsShowText.NoMatch Then
    sb1.SimpleText = "Current tech level is '" + CStr(TechLevel) + "'.   " + rsShowText.Fields("Text").Value
End If
End Sub

