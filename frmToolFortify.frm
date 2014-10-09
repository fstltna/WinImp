VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmToolFortify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fortification Tool"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fortify"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Game Settings"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "67"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Cutoff Level: "
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "127"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "60"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Maximum Moblity: "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Mobilty Gain per Update:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5741
      _Version        =   393216
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmToolFortify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'180206 rjk: Replace ldump with GetLandUnits.

Dim rs As Recordset
Dim MaxMob As Integer
Dim MobPerUpdate As Integer
Dim strSect As String
Dim mobavail As Single
Dim mobused As Single
Dim fortfact As Single
Dim cmd() As String
' Dim ary() As String    8/2003 efj  removed

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim n As Integer
Dim strCmd As String

BuildCommands
frmEmpCmd.SubmitEmpireCommand "bf1", False  'run as batch file
For n = 1 To UBound(cmd, 2)
    strCmd = "fortify " + Left$(cmd(2, n), Len(cmd(2, n)) - 1) + " " + cmd(1, n)
    frmEmpCmd.SubmitEmpireCommand strCmd, True    'sectors
Next n
frmEmpCmd.SubmitEmpireCommand "bf2", False  'run as batch file
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetLandUnits
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Form_Load()
'Set always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
cmdOK.Enabled = False

If Not (rsVersion.BOF And rsVersion.EOF) Then
    MobPerUpdate = rsVersion.Fields("Mob gain - Land")
    MaxMob = rsVersion.Fields("Max mob - Land")
    Set rs = rsLand.Clone
    rs.Index = "PrimaryKey"
    FillGrid
End If
End Sub

Public Sub FillGrid()

' Dim n As Integer    8/2003 efj  removed
' Dim var1 As Variant    8/2003 efj  removed
' Dim totmobavail As Single    8/2003 efj  removed

'Setup grid
grid1.Visible = False
grid1.Clear
grid1.WordWrap = True
grid1.Rows = 2
grid1.Cols = 8
grid1.FixedCols = 1
grid1.RowHeight(0) = -1      'return to default
grid1.RowHeight(0) = grid1.RowHeight(0) * 2

grid1.ColWidth(0) = -1      'return to default
grid1.ColWidth(-1) = grid1.ColWidth(0)
grid1.ColWidth(0) = grid1.ColWidth(0) * 1.5
grid1.ColWidth(1) = grid1.ColWidth(0) / 2#
grid1.ColWidth(2) = grid1.ColWidth(1) / 1.25
grid1.ColWidth(3) = grid1.ColWidth(2)
grid1.ColWidth(4) = grid1.ColWidth(2)
grid1.ColWidth(5) = grid1.ColWidth(2)
grid1.ColWidth(6) = grid1.ColWidth(2)
grid1.ColWidth(7) = grid1.ColWidth(2) * 2.5
grid1.ColAlignment(0) = flexAlignLeftBottom
grid1.ColAlignment(1) = flexAlignCenterBottom
grid1.ColAlignment(2) = flexAlignRightBottom
grid1.ColAlignment(3) = flexAlignRightCenter
grid1.ColAlignment(4) = flexAlignRightCenter
grid1.ColAlignment(5) = flexAlignRightCenter
grid1.ColAlignment(6) = flexAlignRightCenter
grid1.ColAlignment(7) = flexAlignLeftBottom
    
'Fill in row headers
grid1.row = 0
grid1.col = 0
grid1.CellFontBold = True
grid1.Text = "Unit ID"
grid1.col = 1
grid1.CellFontBold = True
grid1.Text = "Loc"
grid1.col = 2
grid1.CellFontBold = True
grid1.Text = "Eff"
grid1.col = 3
grid1.CellFontBold = True
grid1.Text = "Curr Mob"
grid1.col = 4
grid1.CellFontBold = True
grid1.Text = "Curr Fort"
grid1.col = 5
grid1.CellFontBold = True
grid1.Text = "Mob used"
grid1.col = 6
grid1.CellFontBold = True
grid1.Text = "New Fort"
grid1.col = 7
grid1.CellFontBold = True
grid1.Text = "  Action"

If rs.EOF And rs.BOF Then
    Exit Sub
End If

'Find out where the engineers are
strSect = vbNullString
rs.MoveFirst
While Not rs.EOF
    If rs.Fields(1) = "eng" Then
        strSect = strSect + " " + SectorString(rs.Fields("x"), rs.Fields("y")) + " "
    End If
    rs.MoveNext
Wend

rs.MoveFirst
While Not rs.EOF
    grid1.row = grid1.Rows - 1
    grid1.Rows = grid1.Rows + 1

    'Get the unit description
    grid1.col = 0
    grid1.Text = rs.Fields(1) ' Set in case lookup fails
    rsBuild.Seek "=", "l", rs.Fields(1)
    If Not rsBuild.NoMatch Then
        If Len(rsBuild.Fields("desc")) > 0 Then
            grid1.Text = rsBuild.Fields("desc")
        End If
    End If
    grid1.Text = grid1.Text + " #" + CStr(rs.Fields(0))

    grid1.col = 1
    grid1.Text = SectorString(rs.Fields("x"), rs.Fields("y"))
    
    'check for an engineer in the same sector
    If InStr(strSect, grid1.Text) > 0 Then
        fortfact = 1.5
    Else
        fortfact = 1#
    End If
    
    grid1.col = 2
    grid1.Text = CStr(rs.Fields("eff"))
    grid1.col = 3
    grid1.Text = CStr(rs.Fields("mob"))
    grid1.col = 4
    grid1.Text = CStr(rs.Fields("fort"))
    mobused = CInt(0.5 + (MaxMob - rs.Fields("fort")) / fortfact)
    mobavail = rs.Fields("mob") - (MaxMob - MobPerUpdate)
    If mobavail < 0 Then mobavail = 0
    If mobused > mobavail Then mobused = mobavail
    grid1.col = 5
    grid1.Text = CStr(mobused)
    grid1.col = 6
    grid1.Text = (CStr(rs.Fields("fort") + CInt(mobused * fortfact - 0.3)))
    grid1.col = 7
    If MaxMob <= rs.Fields("fort") Then
        grid1.Text = "  Fully Fort."
    ElseIf mobused = 0 Then
        grid1.Text = "  None"
    Else
        grid1.Text = "  Use Excess"
    End If
    rs.MoveNext
Wend

grid1.Visible = True
cmdOK.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub grid1_Click()

' Dim n As Integer    8/2003 efj  removed
Dim id As Integer
' Dim sx As Integer    8/2003 efj  removed
' Dim sy As Integer    8/2003 efj  removed
Dim saverow As Integer

'If you clicked on a header, resort the grid
saverow = grid1.MouseRow
If saverow = 0 Then

'Otherwise, mark the column
Else
    
    grid1.row = saverow
    grid1.col = 0
    id = CInt(Trim$(Mid$(grid1.Text, InStr(grid1.Text, "#") + 1)))
    rs.Seek "=", id
    If rs.NoMatch Then
        Exit Sub
    End If
    
    'check for an engineer in the same sector
    grid1.col = 1
    If InStr(strSect, grid1.Text) > 0 Then
        fortfact = 1.5
    Else
        fortfact = 1#
    End If
    
    grid1.col = 7
    
    'Toggle setting
    Select Case grid1.Text
        Case "  None"
            grid1.Text = "  Use Excess"
            mobavail = rs.Fields("mob") - (MaxMob - MobPerUpdate)
            If mobavail < 0 Then
                grid1.Text = "  Use Req"
            End If
        Case "  Use Excess"
            grid1.Text = "  Use Req"
        Case "  Use Req"
            grid1.Text = "  Use All"
        Case "  Use All"
            grid1.Text = "  None"
    End Select

    'process new one
    Select Case grid1.Text
        Case "  None"
            mobused = 0
        Case "  Use Excess"
            mobused = CInt(0.5 + (MaxMob - rs.Fields("fort")) / fortfact)
            mobavail = rs.Fields("mob") - (MaxMob - MobPerUpdate)
            If mobavail < 0 Then mobavail = 0
            If mobused > mobavail Then mobused = mobavail
        Case "  Use Req"
            mobused = CInt(0.5 + (MaxMob - rs.Fields("fort")) / fortfact)
            mobavail = rs.Fields("mob")
            If mobused >= mobavail Then
                mobused = mobavail
                grid1.Text = "  Use All"
            End If
                    
        Case "  Use All"
            mobused = CInt(0.5 + (MaxMob - rs.Fields("fort")) / fortfact)
            mobavail = rs.Fields("mob")
            If mobused < mobavail Then
                grid1.Text = "  None"
                mobused = 0
            End If
    End Select
    grid1.col = 5
    grid1.Text = CStr(mobused)
    grid1.col = 6
    grid1.Text = CStr(CInt(rs.Fields("fort") + CInt(mobused * fortfact - 0.3)))
End If

End Sub


Public Sub BuildCommands()

Dim n As Integer
Dim m As Integer
Dim found As Boolean


ReDim cmd(2, 0)
For n = 1 To grid1.Rows - 1
    grid1.row = n
    grid1.col = 5
    m = 1
    found = False
    If Val(grid1.Text) > 0 Then
        Do While m <= UBound(cmd, 2)
            If cmd(1, m) = grid1.Text Then
                grid1.col = 0
                cmd(2, m) = cmd(2, m) + Trim$(Mid$(grid1.Text, InStr(grid1.Text, "#") + 1)) + "/"
                found = True
                Exit Do
            End If
            m = m + 1
        Loop
        If Not found Then
            m = UBound(cmd, 2) + 1
            ReDim Preserve cmd(2, m)
            grid1.col = 5
            cmd(1, m) = grid1.Text
            grid1.col = 0
            cmd(2, m) = cmd(2, m) + Trim$(Mid$(grid1.Text, InStr(grid1.Text, "#") + 1)) + "/"
        End If
    End If
Next n

End Sub





