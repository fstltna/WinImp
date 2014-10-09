VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmUnits 
   Caption         =   "Ships"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8400
   Icon            =   "frmUnits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   8400
   Visible         =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   661
      ButtonWidth     =   714
      ButtonHeight    =   503
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ship"
            Object.ToolTipText     =   "Ships"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Land"
            Object.ToolTipText     =   "Lands"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Plane"
            Object.ToolTipText     =   "Planes"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Enemy"
            Object.ToolTipText     =   "Enemy Units"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1700
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1700
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Mob"
            Object.ToolTipText     =   "Only Show if Mobility >0"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eff"
            Object.ToolTipText     =   "Only Show if at least Min. Eff."
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Unit from Database"
            Object.Tag             =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Pin"
            Object.ToolTipText     =   "Pin Unit Box in Place"
            Object.Tag             =   "Pin"
            ImageIndex      =   7
            Style           =   1
         EndProperty
      EndProperty
      MouseIcon       =   "frmUnits.frx":0442
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmUnits.frx":045E
         Left            =   3600
         List            =   "frmUnits.frx":0460
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox cmbSub 
         Height          =   315
         ItemData        =   "frmUnits.frx":0462
         Left            =   1680
         List            =   "frmUnits.frx":0464
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4260
      _Version        =   327680
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":09BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":0F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":1468
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":19BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":1CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":1FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUnits.frx":218C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DisplaySubset As Integer
Public DisplaySelect As Integer
Public SecX As Integer
Public SecY As Integer
Public Fleet As String
Public SubType As Integer

Public FillSubSetFlag As Boolean
Private strSubSect As String
Private strSubFleet As String
Private strSel As String

'Sort Variables
Dim SortCol As Integer
Dim SortDecending As Boolean

'Selection Variables
Dim NeedMob As Boolean
Dim NeedEff As Boolean

Const DSP_SHIP = 1
Const DSP_LAND = 2
Const DSP_PLANE = 3
Const DSP_ENEMY = 4

Const DSP_ALL = 1
Const DSP_FLEET = 2
Const DSP_SECT = 3

Const TYPE_CAPITAL = 1
Const TYPE_ESCORT = 2
Const TYPE_SUPPORTSHIP = 3
Const TYPE_SUB = 4
Const TYPE_CARRIER = 5
Const TYPE_TROOPSHIPS = 6
Const TYPE_MERCHANTS = 7
Const TYPE_FIGHTERS = 8
Const TYPE_BOMBERS = 9
Const TYPE_ESCORTPLANES = 10
Const TYPE_TRANSPORTS = 11
Const TYPE_RECON = 12
Const TYPE_MISSILES = 13
Const TYPE_SATELLITES = 14
Const TYPE_CHOPPER = 15
Const TYPE_ANTISUB = 16
Const TYPE_COMBAT = 17
Const TYPE_ARTILLERY = 18
Const TYPE_SUPPORT = 19
Const TYPE_ENGINEERS = 20

Private strType(1 To TYPE_ENGINEERS) As String

' Calculate index for use with TextArray property.
Function gridIndex(row As Integer, col As Integer) As Long
gridIndex = row * grid1.Cols + col
End Function

Public Sub FillGrid()

Dim rs As Recordset
Dim fieldcnt As Integer
Dim x As Integer
Dim substype As String
Dim substypekey As String
Dim BuildType As String
Dim tstx As Long
Dim tsty As Long

On Error Resume Next

'Reset sort variables
SortCol = -1
SortDecending = True

FillTypeBox

'Determain which type of unit we are displaying
Select Case DisplaySelect
    Case DSP_SHIP
        If SubType = TYPE_MERCHANTS Then
            Set rs = rsShipMerchant
        Else
            Set rs = rsShipList
        End If
        
        rs.Requery
        Caption = "Ships"
        substype = "Fleet"
        substypekey = "flt"
        BuildType = "s"
        
    Case DSP_LAND
        Set rs = rsLandList
        rs.Requery
        Caption = "Land Units"
        substype = "Army"
        substypekey = "army"
        BuildType = "l"
    
    Case DSP_PLANE
        Set rs = rsPlaneList
        rs.Requery
        Caption = "Planes"
        substype = "Wing"
        substypekey = "wing"
        BuildType = "p"
    
    Case DSP_ENEMY
        Set rs = rsEnemyUnit
        rs.Requery
        Caption = "Enemy Units"
        substype = "Nation"
        substypekey = "owner"
End Select

If DisplaySubset = DSP_SECT Then
    Caption = Caption + " in Sector " + CStr(SecX) + "," + CStr(SecY)
ElseIf DisplaySubset = DSP_FLEET Then
    If DisplaySelect = DSP_ENEMY Then
        Caption = Caption + " in " + substype + " " + Fleet
    Else
        Caption = Caption + " owned by " + substype
    End If
End If

'if you are refilling the subset combo box, clear it now
If FillSubSetFlag Then
    cmbSub.Clear
    strSubSect = ""
    strSubFleet = ""
    cmbSub.AddItem "All"
End If

rs.Index = "Primary Key"

fieldcnt = rs.Fields.Count
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.Cols = fieldcnt - 1
grid1.ColWidth(0) = -1      'return to default
grid1.ColWidth(-1) = grid1.ColWidth(0) / 2

If DisplaySelect = DSP_PLANE Or DisplaySelect = DSP_ENEMY Then
    grid1.ColWidth(0) = grid1.ColWidth(0) * 4#
Else
    grid1.ColWidth(0) = grid1.ColWidth(0) * 3.5
End If

If DisplaySelect = DSP_ENEMY Then
    grid1.ColWidth(7) = grid1.ColWidth(0) * 0.8
End If

grid1.row = 0
'Fill in row headers
For x = 2 To fieldcnt - 1
    grid1.col = x - 1
    grid1.Text = rs.Fields(x).Name
    grid1.CellAlignment = flexAlignRightCenter
    grid1.CellFontBold = True
Next x


'Fill in rows
rs.MoveFirst
While Not rs.EOF

    'Add to the sub combo box
    If FillSubSetFlag Then
        AddSubBox substype, rs.Fields(substypekey), rs.Fields("x"), rs.Fields("y")
    End If
    
    'Check to see if we are displaying this fleet
    If (DisplaySubset = DSP_ALL) Or _
       (DisplaySubset = DSP_FLEET And Fleet = Trim(rs.Fields(substypekey))) Or _
       (DisplaySubset = DSP_SECT And SecX = rs.Fields("x") And SecY = rs.Fields("y")) Then
    
        'check to see if it is in the selected type subset
        If SubType = 0 Or InStr(strType(SubType), rs.Fields(1)) Then
        
            'check the mob and eff requirement
            If DisplaySelect = DSP_ENEMY Or _
            (((Not NeedMob) Or rs.Fields("mob") > 0) And _
            ((Not NeedEff) Or rs.Fields("eff") > 59 Or _
            (DisplaySelect = DSP_PLANE And rs.Fields("eff") > 39))) Then
        
                grid1.row = grid1.row + 1
                grid1.col = 0
                
                If DisplaySelect = DSP_ENEMY Then
                    BuildType = rs.Fields(9)    'Check the class
                End If
                
                'Get the unit description
                grid1.Text = rs.Fields(1) ' Set in case lookup fails
                rsBuild.Seek "=", BuildType, rs.Fields(1)
                If Not rsBuild.NoMatch Then
                    If Len(rsBuild.Fields("desc")) > 0 Then
                        grid1.Text = rsBuild.Fields("desc")
                    End If
                End If
                
                grid1.Text = grid1.Text + " #" + CStr(rs.Fields(0))
                grid1.CellAlignment = flexAlignLeftCenter
                
                'grid1.CellFontBold = True
                For x = 2 To fieldcnt - 1
                    grid1.col = x - 1
                    grid1.Text = rs.Fields(x).Value
                    If DisplaySelect = DSP_ENEMY And (grid1.Text = "-1") Then
                        grid1.Text = "???"
                    End If
                    grid1.CellAlignment = flexAlignRightCenter
                Next x
                grid1.Rows = grid1.Rows + 1
            End If
        End If
    End If
    rs.MoveNext
Wend

'Reshow the grid
grid1.Visible = True

'Set the option buttons to the correct entry
Toolbar1.Buttons(DisplaySelect).Value = tbrPressed

'Set the cmbSelect if box was refilled
If FillSubSetFlag Then
    Select Case DisplaySubset
        Case DSP_ALL
            cmbSub.ListIndex = 0
        Case DSP_FLEET
            x = 1
            While Left(cmbSub.List(x), 4) <> Left(substype, 4) Or _
                Right(cmbSub.List(x), 1) <> Fleet
                x = x + 1
            Wend
            If x < cmbSub.ListCount - 1 Then
                cmbSub.ListIndex = x
            End If
        Case DSP_SECT
            x = 1
            Do
                If Left(cmbSub.List(x), 4) = "Sect" Then
                    tstx = cmbSub.ItemData(x) / 10000
                    tsty = cmbSub.ItemData(x) - tstx * 10000
                    If tstx >= 1000 Then
                        tstx = 0 - tstx + 1000
                    End If
                    If tsty >= 1000 Then
                        tsty = 0 - tsty + 1000
                    End If
                    If tstx = SecX And tsty = SecY Then
                        cmbSub.ListIndex = x
                        Exit Do
                    End If
                End If
            
                x = x + 1
            Loop Until x > cmbSub.ListCount - 1
    End Select
End If

'Sub box is now filled
FillSubSetFlag = False
End Sub


Private Sub cmbSub_Click()

Dim strSelect As String
Dim bUpdate As Boolean
Dim n As Integer
Dim n2 As Integer
Dim NewSecX As Integer
Dim NewSecY As Integer

'Boolean keeps track if an update of grid is needed
bUpdate = False

'First check if all was selected
strSelect = cmbSub.List(cmbSub.ListIndex)
If strSelect = "All" Then
    If DisplaySubset <> DSP_ALL Then
        bUpdate = True
        DisplaySubset = DSP_ALL
    End If

'Then See if a sector was selected
ElseIf Left(strSelect, 4) = "Sect" Then
    n = InStr(9, strSelect, ",")
    NewSecX = CInt(Mid(strSelect, 9, n - 9))
    n2 = InStr(n + 1, strSelect, ")")
    NewSecY = CInt(Mid(strSelect, n + 1, n2 - (n + 1)))
    
    'Make sure selection criteria has changed
    If DisplaySubset <> DSP_SECT Then
        bUpdate = True
        DisplaySubset = DSP_SECT
        SecX = NewSecX
        SecY = NewSecY
    Else
        If SecX <> NewSecX Or SecY <> NewSecY Then
            bUpdate = True
            SecX = NewSecX
            SecY = NewSecY
        End If
    End If
Else

' Finally, must be a fleet/army/wing
    If DisplaySubset <> DSP_FLEET Then
        bUpdate = True
        DisplaySubset = DSP_FLEET
        Fleet = Right(strSelect, 1)
    Else
        If Fleet <> Right(strSelect, 1) Then
            Fleet = Right(strSelect, 1)
            bUpdate = True
        End If
    End If
End If

If bUpdate Then
    FillGrid
End If

End Sub


Private Sub AddSubBox(stype As String, Fleet As String, x As Integer, Y As Integer)

Dim strSec As String
Dim n As Long

'If fleet has not been added
If InStr(strSubFleet, Fleet) = 0 Then
    strSubFleet = strSubFleet + Fleet
    cmbSub.AddItem stype + " " + Fleet
End If

strSec = "(" + CStr(x) + "," + CStr(Y) + ")"

If InStr(strSubSect, strSec) = 0 Then
    strSubSect = strSubSect + strSec
    
    'Encode sector numbers in item data
    If x < 0 Then
        n = 0 - x + 1000
    Else
        n = x
    End If
    n = n * 10000
    If Y < 0 Then
        n = n - Y + 1000
    Else
        n = n + Y
    End If
    
    cmbSub.AddItem "Sector " + strSec
    cmbSub.ItemData(cmbSub.NewIndex) = n
End If


End Sub

Private Sub ShrinkForm()

'Don't try to resize when form is minimized
If WindowState <> vbNormal Then
    Exit Sub
End If

'Try to resize the box so that the close button is always visible
Dim n As Long
n = Me.Width
If Me.Left + Me.Width > Screen.Width Then
    n = Screen.Width - Me.Left
End If

If n < Screen.Width / 5 Then
    n = Screen.Width / 5
End If
Me.Width = n

End Sub


Private Sub DeleteUnits()
Dim Startx As Integer
Dim Finishx As Integer
Dim x1 As Integer
Dim gi As Integer
Dim n As Integer
Dim strOwner As String
Dim strClass As String
Dim strID As String
Dim hldIndex As String
Dim rs As Recordset

Startx = grid1.row
Finishx = grid1.RowSel

If Finishx < Startx Then
    x1 = Startx
    Startx = Finishx
    Finishx = x1
End If

'set recordset
Select Case DisplaySelect
    Case DSP_SHIP
        Set rs = rsShip
    Case DSP_LAND
        Set rs = rsLand
    Case DSP_PLANE
        Set rs = rsPlane
    Case DSP_ENEMY
        Set rs = rsEnemyUnit
End Select

'set index
hldIndex = rs.Index
rs.Index = "PrimaryKey"

'go thru selected rows, delete units
For x1 = Startx To Finishx
    gi = gridIndex(x1, 0)
    n = InStr(grid1.TextArray(gi), "#")
    If n > 0 Then
        strID = Trim(Mid$(grid1.TextArray(gi), n + 1))
        If DisplaySelect = DSP_ENEMY Then
            gi = gridIndex(x1, 3)
            strOwner = grid1.TextArray(gi)
            gi = gridIndex(x1, 8)
            strClass = grid1.TextArray(gi)
            rs.Seek "=", CInt(strOwner), strClass, CInt(strID)
        Else
            rs.Seek "=", CInt(strID)
        End If
        If Not rs.NoMatch Then
            rs.Delete
        End If
    End If
Next x1
rs.Index = hldIndex
Set rs = Nothing

'redraw the map
frmDrawMap.DrawHexes
FillGrid

End Sub

Private Sub cmbType_Click()
SubType = cmbType.ItemData(cmbType.ListIndex)
FillGrid
End Sub

Private Sub Form_Activate()
ShrinkForm
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'check and see if the control key is down
If Not ((Shift And vbCtrlMask) = 2) Then
    Exit Sub
End If
On Error Resume Next

If KeyCode = vbKey1 Then
    frmEmpCmd.WindowState = vbNormal
    frmEmpCmd.SetFocus
End If
If KeyCode = vbKey2 Then
    frmUnits.ChangeState vbNormal
    frmTelegram.SetFocus
End If
If KeyCode = vbKey3 Then
    frmTelegram.WindowState = vbNormal
    frmTelegram.SetFocus
End If

End Sub

Private Sub Form_Load()
FillSubSetFlag = True
cmbSub.ZOrder 0
Toolbar1.ZOrder 1
cmbSub.Move Toolbar1.Buttons(5).Left, Toolbar1.Buttons(5).top
cmbType.Move Toolbar1.Buttons(7).Left, Toolbar1.Buttons(7).top

strType(TYPE_CAPITAL) = " bbc bb hc lc mc ncr "
strType(TYPE_ESCORT) = " ad af frg dd mf nas pt "
strType(TYPE_SUPPORTSHIP) = " aac agc ms mb nsp "
strType(TYPE_SUB) = " sb sbc msb na nw"
strType(TYPE_CARRIER) = " car cal can cvl cel cvf "
strType(TYPE_TROOPSHIPS) = " ls tt "
strType(TYPE_MERCHANTS) = " cs fb ft od oe os ss tk "
strType(TYPE_FIGHTERS) = " f1 f2 jf1 jf2 fb jfb np "
strType(TYPE_BOMBERS) = " lb jl mb fb jfb hb jhb sb "
strType(TYPE_ESCORTPLANES) = " es jes "
strType(TYPE_TRANSPORTS) = " tc tr jt "
strType(TYPE_RECON) = " zep sp "
strType(TYPE_MISSILES) = " mi sam ssm srbm irbm icbm slbm asat "
strType(TYPE_SATELLITES) = " lst ss "
strType(TYPE_CHOPPER) = " tc nc ac "
strType(TYPE_ANTISUB) = " as "
strType(TYPE_COMBAT) = " cav inf linf mtif mif lar har arm mar mil reg vet elt gua "
strType(TYPE_ARTILLERY) = " art lat mat hat aau "
strType(TYPE_SUPPORT) = " rad sup sec spy tra "
strType(TYPE_ENGINEERS) = " eng meng elt gua "
  

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Only unload if form Draw says OK
If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
    If Not frmDrawMap.ShutDown Then
        Cancel = 1
        Me.WindowState = vbMinimized
    End If
End If

End Sub

Private Sub Form_Resize()

'Don't try to redraw when form is minimized
If WindowState = vbMinimized Then
    Exit Sub
End If

'Make Stay always on top
Dim success As Long
success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

grid1.top = Toolbar1.Height
'If Me.Width < Toolbar1.Width Then
'    Me.Width = Toolbar1.Width
'End If

Dim h
h = Me.ScaleHeight - grid1.top
If h > 0 Then
    grid1.Height = Me.ScaleHeight - grid1.top
End If
grid1.Width = Me.ScaleWidth

'set combo box
cmbSub.Left = Toolbar1.Buttons(6).Left
cmbSub.top = Toolbar1.Buttons(6).top
cmbSub.Width = Toolbar1.Buttons(6).Width

cmbType.Left = Toolbar1.Buttons(8).Left
cmbType.top = Toolbar1.Buttons(8).top
cmbType.Width = Toolbar1.Buttons(8).Width

End Sub

Private Sub grid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

Dim tb As TextBox
Dim Startx As Integer
Dim Finishx As Integer
Dim x1 As Integer
Dim gi As Integer
Dim n As Integer
Dim bShift As Boolean
Dim bCtl As Boolean
Dim found As Boolean

Dim CurrUnit As Variant
Dim NewUnit As Variant

'If clicking in top row, resort current box
If grid1.MouseRow = 0 Then
    SortGrid grid1.MouseCol
    Exit Sub
End If

'this routine is used to fill prompt boxes
If Not frmDrawMap.PromptUp Then
    Exit Sub
End If

'See if this is a box you add units to
Set tb = frmDrawMap.PromptForm.ActiveControl
If Not (tb.Name = "txtUnit" Or tb.Name = "txtUnit2" _
        Or tb.Name = "txtEscort" Or tb.Name = "txtUnit3" Or _
        (tb.Name = "txtTarget" And DisplaySelect = DSP_ENEMY)) Then
    Exit Sub
End If

bShift = False
If (Shift And vbShiftMask) = 1 Then
    bShift = True
End If

bCtl = False
If (Shift And vbCtrlMask) = 2 Then
    bCtl = True
End If

'If control key is down, clear existing string
If bCtl Then
    tb.Text = ""
End If

'Get current units
If Len(tb.Text) > 0 Then
    CurrUnit = ParseUnitString(tb.Text)
Else
    ReDim CurrUnit(0)
    CurrUnit(0) = ""
End If

Startx = grid1.row
Finishx = grid1.RowSel

If Finishx < Startx Then
    x1 = Startx
    Startx = Finishx
    Finishx = x1
End If

strSel = ""

'go thru selected rows, and add units to selet String
For x1 = Startx To Finishx
    gi = gridIndex(x1, 0)
    n = InStr(grid1.TextArray(gi), "#")
    If n > 0 Then
        strSel = Trim(Mid$(grid1.TextArray(gi), n + 1))
        found = False
        For n = LBound(CurrUnit) To UBound(CurrUnit)
            If CurrUnit(n) = strSel Then
                found = True
            End If
        Next n
        If Not found Then
            If Len(tb.Text) > 0 Then
                tb.Text = tb.Text + "/"
            End If
            tb.Text = tb.Text + strSel
        End If
    End If
Next x1

frmDrawMap.PromptForm.SetFocus

End Sub

Private Sub SortGrid(col As Integer)

'This does a simple bubble sort of the grid in place.  Due to the small size
'of the sample, a more efficent sort is usually not necessary.

Dim n As Integer
Dim row As Integer
Dim var1 As Variant
Dim var2 As Variant

'If clicked multiple times, change the sort order
If SortCol = col Then
    SortDecending = Not SortDecending
End If
SortCol = col

grid1.Visible = False

n = grid1.Rows - 3
While n > 0
    For row = 1 To n
        var1 = grid1.TextArray(gridIndex(row, col))
        var2 = grid1.TextArray(gridIndex(row + 1, col))
        
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

grid1.Visible = True

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

Dim n As Integer
Select Case Button.Key
    Case "Ship"
        n = DSP_SHIP
    Case "Land"
        n = DSP_LAND
    Case "Plane"
        n = DSP_PLANE
    Case "Enemy"
        n = DSP_ENEMY
    Case "Mob"
        NeedMob = (Button.Value = tbrPressed)
        FillGrid
        Exit Sub

    Case "Delete"
        DeleteUnits
        Exit Sub
    
    Case "Pin"
        Exit Sub

    Case "Eff"
        NeedEff = (Button.Value = tbrPressed)
        FillGrid
        Exit Sub
End Select

FillTypeBox

'If option Changed - Refill the grid
If DisplaySelect <> n Then
    DisplaySelect = n
    
    'if showing fleet/wing/army, shift to all
    If DisplaySubset = DSP_FLEET Then
        DisplaySubset = DSP_ALL
    End If

    'Refill
    FillSubSetFlag = True
    FillGrid
End If

End Sub

Private Sub FillTypeBox()

Static CurrentType As Integer
Dim n As Integer

If CurrentType <> DisplaySelect Then
    CurrentType = DisplaySelect
    cmbType.Clear
    cmbType.AddItem "All"
    Select Case DisplaySelect
        Case DSP_SHIP
            cmbType.AddItem "Capital Ships"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_CAPITAL
            cmbType.AddItem "Escorts"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_ESCORT
            cmbType.AddItem "Subs"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_SUB
            cmbType.AddItem "Support"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_SUPPORTSHIP
            cmbType.AddItem "Carriers"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_CARRIER
            cmbType.AddItem "Transports"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_TROOPSHIPS
            cmbType.AddItem "Merchants"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_MERCHANTS
        
        Case DSP_LAND
            cmbType.AddItem "Combat"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_COMBAT
            cmbType.AddItem "Artillery"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_ARTILLERY
            cmbType.AddItem "Support"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_SUPPORT
            cmbType.AddItem "Engineers"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_ENGINEERS
        
        Case DSP_PLANE
            cmbType.AddItem "Fighters"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_FIGHTERS
            cmbType.AddItem "Bombers"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_BOMBERS
            cmbType.AddItem "Escorts"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_ESCORTPLANES
            cmbType.AddItem "Cargo"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_TRANSPORTS
            cmbType.AddItem "Recon"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_RECON
            cmbType.AddItem "Missiles"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_MISSILES
            cmbType.AddItem "Satellites"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_SATELLITES
            cmbType.AddItem "Choppers"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_CHOPPER
            cmbType.AddItem "Sub Hunters"
            cmbType.ItemData(cmbType.NewIndex) = TYPE_ANTISUB
        Case DSP_ENEMY
    End Select
    cmbType.ListIndex = 0
End If

'Make sure there are Type records
If cmbType.ListIndex < 0 Then
    Exit Sub
End If

If SubType <> cmbType.ItemData(cmbType.ListIndex) Then
    For n = 0 To cmbType.ListCount - 1
        If SubType = cmbType.ItemData(n) Then
            cmbType.ListIndex = n
        End If
    Next n
End If

End Sub

'This was added to support the pin button. Some people use wide
'screens and didn't like the unit box popping up and down

Public Sub ChangeState(NewState As Integer)

'If the pin button is not down, go ahead and change it
If Toolbar1.Buttons("Pin").Value = tbrUnpressed Then
    Me.WindowState = NewState
    Exit Sub
End If

'Alway bring it up if thats the request
If Me.WindowState = vbMinimized And NewState <> vbMinimized Then
    Me.WindowState = NewState
End If

'Otherwise, ignore the request

End Sub
'This was added to support the pin button. Some people use wide
'screens and didn't like the unit box popping up and down

Public Sub MoveMe(nleft As Single, ntop As Single, Optional nwidth As Single, Optional nheight As Single)

'If the pin button is down, don't move it
If Toolbar1.Buttons("Pin").Value = tbrPressed Then
    Exit Sub
End If

'fill in optional parms
If nwidth = 0 Then
    nwidth = Me.Width
End If

If nheight = 0 Then
    nheight = Me.Height
End If

If ntop < 0 Then
    ntop = 0
End If

If nleft < 0 Then
    nleft = 0
End If

'Move the box
Me.Move nleft, ntop, nwidth, nheight

End Sub

