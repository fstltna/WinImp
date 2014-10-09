VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmToolFeeder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Famine Relief Tool"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   22
      Text            =   "10"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CheckBox CheckSetThresh 
      Caption         =   "Increase Food Thresholds after moving food"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Value           =   1  'Checked
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Check Starvation"
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Distribute Food"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Food Level "
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "Insure Growth"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Avoid Starvation"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Projected Resources Used"
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
      Begin VB.Label Label2 
         Caption         =   "Shortage: "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "240 food, 27.4 mobility"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resources Limits"
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "range"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "food"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "mob"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4260
      _Version        =   393216
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label5 
      Caption         =   "Percent"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Set Threshold higher than required by"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmToolFeeder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'171003 rjk: Added Strength fields to Sector database
'201103 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim PackDes As String
Dim rs As Recordset
Dim InSetup As Boolean
Dim SortCol As Integer
Dim SortDecending As Boolean
Dim SectorMob As Integer
Dim SectorFood As Integer

'drk 5/25/03:  made general to work with more commodities than just food (e.g. maybe we want to move civs from areas of overpopulation to areas of shortage)
Dim comm_name As String
Dim comm_dist As String

Private Type record  'drk 5/24/03 arrays of variants suck squid.  Removed that and added a proper user defined type
    secx As Integer
    secy As Integer
    comm_reqd As Variant
    comm_shortage As Integer
    mcost As Single
    pcost As Single
    checked As Boolean
End Type

Dim aryValues() As record

Private Sub CheckSetThresh_Click()
    Label5.Enabled = CheckSetThresh
    Label4.Enabled = CheckSetThresh
    Text2.Enabled = CheckSetThresh
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim n As Integer
Dim var1 As record
Dim mobavail As Single
Dim foodavail As Integer
Dim strCmd As String

mobavail = SectorMob - Val(Text1(0).Text)
foodavail = SectorFood - Val(Text1(1).Text)

frmEmpCmd.SubmitEmpireCommand "bf1", False  'run as batch file
For n = 1 To UBound(aryValues)
    var1 = aryValues(n)
    If var1.checked = True Then
        If mobavail > var1.mcost And foodavail > var1.comm_shortage Then
            foodavail = foodavail - var1.comm_shortage
            mobavail = mobavail - var1.mcost
            strCmd = "move " + comm_name + " " + txtOrigin.Text + " " + CStr(var1.comm_shortage) _
                + " " + SectorString(var1.secx, var1.secy)
            frmEmpCmd.SubmitEmpireCommand strCmd, False    'sectors
                 
            'drk@unxsoft.com 6/23/02
            If CheckSetThresh Then
                Dim newthresh As Integer
                newthresh = Int(0.5 + (var1.comm_reqd * (1 + Val(Text2) / 100)))
                rsSectors.Seek "=", var1.secx, var1.secy
                'drk 7/20/02: this can get very BTU intensive for large countries.  Save some by only submitting the thresh command if we're actually increasing the thresh
                If (newthresh > rsSectors.Fields(comm_dist)) Then
                    strCmd = "thresh " + comm_name + " " & SectorString(var1.secx, var1.secy) & " " & newthresh
                    frmEmpCmd.SubmitEmpireCommand strCmd, False
                End If
            End If
            
        End If
    End If
Next n
frmEmpCmd.SubmitEmpireCommand "bf2", False  'run as batch file
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Command1_Click()
frmEmpCmd.SubmitEmpireCommand "bf1", False  'run as batch file
frmEmpCmd.SubmitEmpireCommand "starvation *", True    'sectors
frmEmpCmd.SubmitEmpireCommand "bf2", False  'run as batch file
frmEmpCmd.SubmitEmpireCommand "db2", False  'force map redraw
End Sub

Private Sub Form_Load()
'Set always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
InSetup = True

lblError.Caption = vbNullString
Label2(0).Caption = vbNullString
Label2(1).Caption = vbNullString
Text1(2) = 99
Set rs = rsSectors.Clone
rs.Index = "PrimaryKey"
SortDecending = True
SortCol = 2
ReDim aryValues(0)

comm_name = "food"
comm_dist = "f_dist"
End Sub

Public Sub FillGrid(ClearFirst As Boolean)

Dim n As Integer
Dim var1 As record
Dim totmobavail As Single
Dim mobavail As Single
Dim mobused As Single
Dim mobtotal As Single
Dim totfoodavail As Integer
Dim foodavail As Integer
Dim foodused As Integer
Dim foodtotal As Integer
Dim labels As Variant

If InSetup Then
    Exit Sub
End If

'Setup grid
grid1.Visible = False
If ClearFirst Then
    grid1.Clear
    
    grid1.Rows = 2
    
    grid1.Cols = 7
    grid1.FixedCols = 2
    grid1.ColWidth(0) = -1      'return to default
    grid1.ColWidth(0) = grid1.ColWidth(0) / 2.5
    grid1.ColWidth(1) = grid1.ColWidth(0)
    grid1.ColWidth(2) = grid1.ColWidth(0) * 2.1
    grid1.ColWidth(3) = grid1.ColWidth(0) * 2.1
    grid1.ColWidth(6) = grid1.ColWidth(0) * 1.9
    grid1.ColAlignment(0) = flexAlignCenterCenter
    grid1.ColAlignment(1) = flexAlignCenterCenter
    grid1.ColAlignment(2) = flexAlignRightCenter
    grid1.ColAlignment(3) = flexAlignRightCenter
    grid1.ColAlignment(4) = flexAlignRightCenter
    grid1.ColAlignment(6) = flexAlignCenterCenter
    
    'Fill in row headers
    grid1.col = 0
    grid1.row = 0
    
    labels = Array("X", "Y", comm_name + " req.", "shortage", "path cost", "mob. cost", "Select")
    For n = 0 To 6
        grid1.col = n
        grid1.CellFontBold = True
        grid1.Text = labels(n)
    Next n
    
End If

totmobavail = SectorMob - Val(Text1(0).Text)
totfoodavail = SectorFood - Val(Text1(1).Text)
mobavail = totmobavail
foodavail = totfoodavail
mobused = 0
foodused = 0

For n = 1 To UBound(aryValues)
    var1 = aryValues(n)
   
    If ClearFirst Then
        grid1.row = grid1.Rows - 1
        grid1.Rows = grid1.Rows + 1
    Else
        grid1.row = n
    End If

    grid1.col = 0
    grid1.Text = CStr(record_nth_elt(var1, grid1.col))
    grid1.col = 1
    grid1.Text = CStr(record_nth_elt(var1, grid1.col))
    
    'Set the check value
    grid1.col = 6 '
    If var1.checked = True Then '
        grid1.CellForeColor = vbBlack '
        grid1.Text = "Yes" '
        
        mobtotal = mobtotal + var1.mcost
        foodtotal = foodtotal + var1.comm_shortage
        If mobavail > var1.mcost And foodavail > var1.comm_shortage Then
            mobused = mobused + var1.mcost
            foodused = foodused + var1.comm_shortage
        End If
       
        grid1.col = 2
        grid1.CellForeColor = vbBlack
        grid1.Text = CStr(record_nth_elt(var1, grid1.col)) + "   "
       
        grid1.col = 3
        If foodavail > var1.comm_reqd Then
            foodavail = foodavail - var1.comm_reqd
            grid1.CellForeColor = vbBlack
        Else
            grid1.CellForeColor = vbRed
        End If
        grid1.Text = CStr(record_nth_elt(var1, grid1.col)) + "   "
        
        
        grid1.col = 4
        grid1.CellForeColor = vbBlack
        grid1.Text = Format$(CSng(record_nth_elt(var1, grid1.col)), "###0.0000") + "   "
        grid1.col = 5
        If mobavail > var1.mcost Then
            mobavail = mobavail - var1.mcost
            grid1.CellForeColor = vbBlack
        Else
            grid1.CellForeColor = vbRed
        End If
        grid1.Text = Format$(CSng(record_nth_elt(var1, grid1.col)), "###0.0000") + "   "
    Else
        grid1.col = 2
        grid1.CellForeColor = vbBlue
        grid1.Text = CStr(record_nth_elt(var1, grid1.col)) + "   "
        grid1.col = 3
        grid1.CellForeColor = vbBlue
        grid1.Text = CStr(record_nth_elt(var1, grid1.col)) + "   "
        grid1.col = 4
        grid1.CellForeColor = vbBlue
        grid1.Text = Format$(CSng(record_nth_elt(var1, grid1.col)), "###0.0000") + "   "
        grid1.col = 5
        grid1.CellForeColor = vbBlue
        grid1.Text = Format$(CSng(record_nth_elt(var1, grid1.col)), "###0.0000") + "   "
        grid1.col = 6
        grid1.CellForeColor = vbBlue
        grid1.Text = "No"
   End If
    
Next n

Label2(0) = CStr(foodused) + " " + comm_name + ", " + CStr(Int(mobused)) + " mob."
Label2(1) = "Shortage: "
If foodtotal > totfoodavail Then
    Label2(1) = Label2(1) + CStr(foodtotal - totfoodavail) + " " + comm_name + " "
End If
If mobtotal > totmobavail Then
    Label2(1) = Label2(1) + CStr(Int(mobtotal - totmobavail)) + " mob."
End If

If Len(Label2(1)) = 10 Then
    Label2(1) = Label2(1) + "None"
End If

grid1.Visible = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub grid1_Click()

Dim n As Integer
Dim var1 As record
Dim sx As Integer
Dim sy As Integer
' Dim saverow As Integer    8/2003 efj  removed

'If you clicked on a header, resort the grid
If grid1.MouseRow = 0 Then
    If SortCol = grid1.MouseCol Then
        SortDecending = Not SortDecending
    End If
    SortCol = grid1.MouseCol
    SortValues
    FillGrid False

'Otherwise, mark the column
ElseIf grid1.MouseRow < grid1.Rows Then 'drk@unxsoft.com, bugfix. would crash on certain pathological conditions
    If grid1.TextMatrix(grid1.MouseRow, 0) <> vbNullString Then 'ditto
'        saverow = grid1.MouseRow    8/2003 efj  removed
        grid1.row = grid1.MouseRow
        grid1.col = 0
        sx = CInt(grid1.Text)
        grid1.col = 1
        sy = CInt(grid1.Text)
        
        grid1.col = 6
        For n = 1 To UBound(aryValues)
            var1 = aryValues(n)
            If sx = var1.secx And sy = var1.secy Then
                If grid1.Text = "Yes" Then
                    grid1.Text = "No"
                    var1.checked = 0
                Else
                    grid1.Text = "Yes"
                    var1.checked = 1
                End If
                aryValues(n) = var1
            End If
        Next n
        
        FillGrid False
    End If
End If
End Sub

Private Sub Option1_Click(Index As Integer)
LoadValues
SortValues
FillGrid True
End Sub

Private Sub Text1_LostFocus(Index As Integer)
LoadValues
SortValues
FillGrid True
End Sub

Private Sub txtOrigin_Change()

Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, txtOrigin.Text) Then
    FillSectorInfo sx, sy
    InSetup = False
    LoadValues
    SortValues
    FillGrid True
Else
    HandleError "Not a valid Sector"
End If

End Sub


Private Sub FillSectorInfo(secx As Integer, secy As Integer)

'get sector record
rs.Seek "=", secx, secy
If rs.NoMatch Then
    HandleError "Sector not found"
    Exit Sub
End If

SectorMob = rs.Fields("mob")
SectorFood = rs.Fields(comm_name)

lblError.Caption = CStr(SectorFood) + " " + comm_name + ", " _
    + CStr(SectorMob) + " mob."

'set sector type
If rs.Fields("eff") < 60 Then
    PackDes = vbNullString
Else
    PackDes = rs.Fields("des")
End If

Text1(0) = Int(SectorMob / 2)
Text1(1) = FoodRequired(rs)

End Sub
Private Sub HandleError(ErrMsg As String)
lblError.Caption = ErrMsg
Text1(0).Text = vbNullString
Text1(1).Text = vbNullString
Text1(2).Text = vbNullString
Label2(0).Caption = vbNullString
Label2(1).Caption = vbNullString
grid1.Clear
End Sub
Private Function PathCost() As Single

On Error Resume Next
Dim pvar As Variant
'Dim pcost As Single 'drk 5/24/03 duh

'Compute the path cost
pvar = BestPath(txtOrigin.Text, SectorString(rs.Fields("x").Value, rs.Fields("y").Value), MT_COMMODITY)
'pcost = pvar(2)
If Len(CStr(pvar(1))) = 0 Then
    PathCost = 0
    Exit Function
End If

'Compute the mobility cost
'PathCost = pcost
PathCost = pvar(2)

End Function

Private Sub LoadValues()

Dim StartX As Integer
Dim StartY As Integer
Dim MaxRange As Integer
Dim n As Integer
Dim fval As record
Dim StrMsg As String

If InSetup Then
    Exit Sub
End If

ReDim aryValues(0)

If Not ParseSectors(StartX, StartY, txtOrigin.Text) Then
    Exit Sub
End If

MaxRange = Val(Text1(2).Text)
If MaxRange <= 0 Then
    Exit Sub
End If

n = 1
'Run through all the sector checking food requirements
rs.MoveFirst
While Not rs.EOF
    If SectorDistance(rs.Fields("x").Value, rs.Fields("y").Value, StartX, StartY) <= MaxRange Then
        If comm_name = "food" Then
        
            fval.comm_reqd = FoodRequired(rs)
            
            If Option1(1).Value Then
                'base on calculated food
                fval.comm_shortage = FoodRequired(rs) - rs.Fields("food")
            Else
                'base on Starvation event markers
                fval.comm_shortage = 0
                StrMsg = EventMarkers.Find(rs.Fields("x").Value, rs.Fields("y").Value)
                If InStr(StrMsg, "Starvation") > 0 Then
                    fval.comm_shortage = Val(StringBetween(StrMsg, "p/", "f"))
                End If
            End If
        ElseIf comm_name = "civ" Then
            fval.comm_reqd = 768 'fixme
            
        End If
        
        If fval.comm_shortage > 0 Then
            fval.pcost = PathCost
            If fval.pcost > 0 Then
                fval.mcost = fval.pcost * MobilityWeight("f", fval.comm_shortage, PackDes)
                'save the values
                fval.secx = rs.Fields("x").Value
                fval.secy = rs.Fields("y").Value
                fval.checked = True
                ReDim Preserve aryValues(n)
                aryValues(n) = fval
                n = n + 1
            End If
        End If
    End If
    rs.MoveNext
Wend

End Sub


Private Sub SortValues()

'This does a simple bubble sort of the array in place.  Due to the small size
'of the sample, a more efficent sort is usually not necessary.

Dim n As Integer
Dim row As Integer
Dim var1 As record
Dim var2 As record
Dim var1c As Single
Dim var2c As Single

If InSetup Then
    Exit Sub
End If

n = UBound(aryValues) - 1
While n > 0
    For row = 1 To n
        var1 = aryValues(row)
        var1c = CSng(record_nth_elt(var1, SortCol))
        var2 = aryValues(row + 1)
        var2c = CSng(record_nth_elt(var2, SortCol))
        
        If SortDecending Then
            If var1c > var2c Then
                aryValues(row) = var2
                aryValues(row + 1) = var1
            End If
        Else
            If var1c < var2c Then
                aryValues(row) = var2
                aryValues(row + 1) = var1
            End If
        End If
    Next row
    n = n - 1
Wend

End Sub

'drk 5/24/03: previously, we had stored things in variant arrays, so it was easy to pick out
'the nth element.  It is preferable to store things in user-defined types, but now we
'need a helper function to pull out the nth element
Private Function record_nth_elt(r As record, S As Integer) As Variant
Select Case S
    Case 0
        record_nth_elt = r.secx
    Case 1
        record_nth_elt = r.secy
    Case 2
        record_nth_elt = r.comm_reqd
    Case 3
        record_nth_elt = r.comm_shortage
    Case 4
        record_nth_elt = r.pcost
    Case 5
        record_nth_elt = r.mcost
    Case 6
        record_nth_elt = r.checked
    Case Else
        Debug.Assert False
End Select
End Function
