VERSION 5.00
Begin VB.Form frmPromptBuild 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2328
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7356
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2328
   ScaleWidth      =   7356
   StartUpPosition =   3  'Windows Default
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
      Left            =   6960
      TabIndex        =   27
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox txtType 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1560
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Builds"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame frmDir 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton cmdDir 
         Caption         =   "N"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "B"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "J"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "H"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "G"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "U"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Y"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtTechLevel 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   1935
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option8 
         Caption         =   "bridge tower"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1515
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "bridge span"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "nuke"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1005
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "plane"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   750
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "land"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   495
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ship"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      Caption         =   "number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Build"
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
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tech Level (Optional)"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmPromptBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092803 rjk: Added timestamp ndump. General reformatting.
'            Added initial field selection.
'            Options now done based on the origin selected.
'100103 rjk: Adjust positions of the options so the text boxes do not overlap.
'100203 rjk: Fixed if a harbor is selected it will have the combo filled with ships
'            (disabled option1 on the form).
'100503 rjk: Removed timestamp ndump.
'101403 rjk: Added ndump timestamp (try #2)
'            Changed initial field selection to be the either direction or combo box.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'080304 rjk: Fixed crash when selecting a sector while building bridge span
'080304 rjk: Changed to determine direction to build from sector selection on the map
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'280804 rjk: For 2K4, explore with m instead c when building bridge spans or towers.
'180206 rjk: Replace ldump, sdump, ndump, lost with GetLandUnits, GetShips, GetNukes and GetLost.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'270407 rjk: Do not automatically select bridge tower when bridge span is available.

Public strCmd As String
Private bFirstLoad As Boolean

Private Sub cmdCancel_Click()
strCmd = vbNullString
Unload Me
End Sub

Private Sub cmdDir_Click(Index As Integer)
txtNum.Text = LCase$(cmdDir(Index).Caption)
cmdOK.Value = True
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim X As Integer
On Error Resume Next
' Dim strCmd2 As String     removed 8/03 ej

'Start with first list box results
If Option1.Value Then
    strCmd = "land"
    frmEmpCmd.HaveLands = True
ElseIf Option2.Value Then
    strCmd = "plane"
    frmEmpCmd.HavePlanes = True
ElseIf Option3.Value Then
    strCmd = "ship"
    frmEmpCmd.HaveShips = True
ElseIf Option4.Value Then
    strCmd = "nuke"
    frmEmpCmd.HaveNukes = True
ElseIf Option7.Value Then
    strCmd = "bridge"
Else
    strCmd = "tower"
End If

strCmd = "build " + strCmd + " " + txtOrigin.Text
If Option7 Or Option8 Then      'if  bridge then fill in
    If Len(Trim$(txtNum.Text)) > 0 Then
        strCmd = strCmd + " " + txtNum.Text
    End If
    frmEmpCmd.SubmitEmpireCommand "bf1", False
Else
    strCmd = strCmd + " " + Trim$(Left$(txtType.Text, 5)) + " "
    'Add number if input
    If Len(txtNum.Text) > 0 Then
        X = Val(txtNum.Text)
        If X < 1 Then
            MsgBox "Number String not valid", vbExclamation + vbOKOnly, "Entry Error"
            txtNum.SetFocus
            Exit Sub
        End If
        strCmd = strCmd + " " + CStr(Val(txtNum.Text))
    End If
    
    'Add tech level if input
    If Len(txtTechLevel.Text) > 0 Then
        If Trim$(txtTechLevel.Text) <> "0" And Val(txtTechLevel.Text) < 1 Then
            MsgBox "Tech String not valid", vbExclamation + vbOKOnly, "Entry Error"
            txtTechLevel.SetFocus
            Exit Sub
        End If
        strCmd = strCmd + " " + CStr(Val(txtTechLevel.Text))
    End If

End If
frmEmpCmd.SubmitEmpireCommand strCmd, True

'If we just built a bridge, try to explore it
If Option7.Value Or Option8.Value Then
    If Len(Trim$(txtNum.Text)) > 0 Then
        If frmOptions.b2K4 Then
            frmEmpCmd.SubmitEmpireCommand "explore m " + txtOrigin.Text + " 1 " + txtNum.Text + "h", True
        Else
            frmEmpCmd.SubmitEmpireCommand "explore c " + txtOrigin.Text + " 1 " + txtNum.Text + "h", True
        End If
    End If
    frmEmpCmd.SubmitEmpireCommand "bf2", False
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If Option3.Value Then
    GetShips
ElseIf Option1.Value Then
    GetLandUnits
ElseIf Option2.Value Then
    GetPlanes
ElseIf Option4.Value Then
    GetNukes
End If
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub cmdShow_Click()
frmToolShow.Show
End Sub

Private Sub Form_Activate()
'txtOrigin.SetFocus
If bFirstLoad Then
    If Option8.Value Or Option7.Value Then
        txtNum.SetFocus
    Else
        txtType.SetFocus
    End If
Else
    '080304 rjk: Fixed crash when selecting a sector while building bridge span
    '080304 rjk: Changed to determine direction to build from sector selection on the map
    If Option8.Value Or Option7.Value Then
        If txtOrigin <> "" Then
            Dim sx As Integer
            Dim sy As Integer
            Dim strPath As String
            
            If ParseSectors(sx, sy, txtOrigin) Then
                If frmDrawMap.SelX <> sx Or frmDrawMap.SelY <> sy Then
                    rsBmap.Seek "=", frmDrawMap.SelX, frmDrawMap.SelY
                    If Not rsBmap.NoMatch Then
                        If rsBmap.Fields("des") = "." Then
                            strPath = EmpirePathDirection(frmDrawMap.SelX - sx, frmDrawMap.SelY - sy)
                            If Len(strPath) = 1 And strPath <> "h" Then
                                txtNum = strPath
                            Else
                                txtNum = ""
                            End If
                        Else
                            txtNum = ""
                        End If
                    Else
                        txtNum = ""
                    End If
                End If
            End If
        End If
    End If
End If
bFirstLoad = False
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
'Option3_Click 092803 rjk: Options now done based on the origin selected.
'101403 rjk: Added bFirstLoad to control what the initial field is
bFirstLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
'Option1 is Lands
Label4.Width = Label3.Width
Label4.Caption = "type"
txtType.Visible = True
txtType.Left = txtOrigin.Left

Label1.Visible = True
Label5.Visible = True
txtNum.Visible = True
txtTechLevel.Visible = True
txtNum.Move txtTechLevel.Left, Label5.top
frmDir.Visible = False

LoadTypebox "l"
End Sub

Private Sub Option2_Click()
Label4.Width = Label3.Width
Label4.Caption = "type"
txtType.Visible = True
txtType.Left = txtOrigin.Left
Label1.Visible = True
Label5.Visible = True
txtNum.Visible = True
txtTechLevel.Visible = True
txtNum.Move txtTechLevel.Left, Label5.top
frmDir.Visible = False

LoadTypebox "p"
End Sub

Private Sub Option3_Click()
Label4.Width = Label3.Width
Label4.Caption = "type"
txtType.Visible = True
txtType.Left = txtOrigin.Left
Label1.Visible = True
Label5.Visible = True
txtNum.Visible = True
txtTechLevel.Visible = True
txtNum.Move txtTechLevel.Left, Label5.top
frmDir.Visible = False

LoadTypebox "s"
End Sub

Private Sub Option4_Click()
Label4.Width = Label3.Width
Label4.Caption = "type"
txtType.Visible = True
txtType.Left = txtOrigin.Left
Label1.Visible = True
Label5.Visible = True
txtNum.Visible = True
txtTechLevel.Visible = True
txtNum.Move txtTechLevel.Left, Label5.top
frmDir.Visible = False

LoadTypebox "n"
End Sub

Private Sub Option7_Click()
Label4.Width = Label1.Width
Label4.Caption = "Direction"
txtType.Visible = False

Label1.Visible = False
Label5.Visible = False
txtNum.Visible = True
txtTechLevel.Visible = False

txtNum.Move txtOrigin.Left + txtOrigin.Width - txtNum.Width, Label4.top
frmDir.Visible = True
frmDir.Move Label4.Left, txtNum.top + txtNum.Height * 1.5
End Sub

Private Sub Option8_Click()
Label4.Width = Label1.Width
Label4.Caption = "Direction"
txtType.Visible = False

Label1.Visible = False
Label5.Visible = False
txtNum.Visible = True
txtTechLevel.Visible = False

txtNum.Move txtOrigin.Left + txtOrigin.Width - txtNum.Width, Label4.top
frmDir.Visible = True
frmDir.Move Label4.Left, txtNum.top + txtNum.Height * 1.5
End Sub

Private Sub Timer1_Timer()
If frmReport.Visible Then
    If frmReport.WindowState = vbNormal Then
        
        'Size Report
        If frmReport.Height > Me.top Then
            frmReport.Height = Me.top
        End If
        
        'move report and shut off timer
        frmReport.top = Me.top - frmReport.Height
        frmReport.Left = Me.Left
        Timer1.Interval = 0
    End If
End If
End Sub

Private Sub LoadTypebox(UnitClass As String)
txtType.Clear
txtType.Text = vbNullString

If Not (rsBuild.BOF And rsBuild.EOF) Then
    rsBuild.MoveFirst
    While Not rsBuild.EOF
        If rsBuild.Fields("type") = UnitClass Then
            If TechLevel < 0 Or rsBuild.Fields("tech") <= TechLevel Then
                txtType.AddItem Left$(rsBuild.Fields("id") + "     ", 5) _
                    + rsBuild.Fields("desc")
            End If
        End If
        rsBuild.MoveNext
    Wend
End If
End Sub

Private Sub SelectOptions() ' 092803 rjk: Options now based on the sector selected.
Dim secx As Integer
Dim secy As Integer

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option7.Enabled = False
Option8.Enabled = False

If ParseSectors(secx, secy, txtOrigin.Text) Then
    rsSectors.Seek "=", secx, secy
    If Not rsSectors.NoMatch Then
        If rsSectors.Fields("coast") = 1 Then
            If Not Option7.Enabled Then
                Option7.Enabled = True
            End If
        Else
            If Option7.Enabled Then
                Option7.Enabled = False
            End If
        End If
        Select Case rsSectors.Fields("des")
        Case "h"
            Option3 = True
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = True
            Option4.Enabled = False
            Option8.Enabled = False
        Case "*"
            Option2 = True
            Option1.Enabled = False
            Option2.Enabled = True
            Option3.Enabled = False
            Option4.Enabled = False
            Option8.Enabled = False
        Case "!"
            Option1 = True
            Option1.Enabled = True
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            Option8.Enabled = False
        Case "n"
            Option4 = True
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = True
            Option8.Enabled = False
        Case "="
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            Option8.Enabled = True
            If Not Option7.Enabled Then
                Option8 = True
            Else
                Option7 = False
                Option8 = False
            End If
        Case Else
            If Option7.Enabled Then
                Option7 = True
            End If
        End Select
    End If
ElseIf txtOrigin.Text <> "" Then
    Option7 = True
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
    Option4.Enabled = True
    Option8.Enabled = True
    Option7.Enabled = True
End If
End Sub

Private Sub txtOrigin_Change()
SelectOptions
End Sub
