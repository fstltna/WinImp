VERSION 5.00
Begin VB.Form frmToolHarborQueue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Build Queue"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Projected Resources Used"
      Height          =   855
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
      Begin VB.Label Label2 
         Caption         =   "99 lcms, 99 hcms, 99 guns, 999 avail"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Cost: $1,400"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add to Queue"
      Height          =   1935
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox txtType 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Number"
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
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double Click to remove"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   2655
   End
End
Attribute VB_Name = "frmToolHarborQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'171003 rjk: Added Strength fields to Sector database
'161103 rjk: Changed the name TxtSector to strSector, code cleanup.
'201103 rjk: Added option to control strength updates
'180304 rjk: if the build list is empty, do not attempt a build, saving generating a server problem.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'270904 rjk: Fixed the crash if the build list is empty.
'261105 rjk: Do not compute avail use the value from the server.
'180206 rjk: Replace sdump, ldump, and pdump with GetShips, GetLandUnits and GetPlanes.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim BuildType As String
Public strSector As String

Private Sub cmdAdd_Click()
Dim X As Integer
Dim i As Integer
Static c As Integer
'check number if input
If Len(txtNum.Text) > 0 Then
    X = Val(txtNum.Text)
    If X < 1 Then
        MsgBox "Number String not valid", vbExclamation + vbOKOnly, "Entry Error"
        txtNum.SetFocus
        Exit Sub
    End If
Else
    X = 1
End If

For i = 1 To X
    c = c + 1
    List1.AddItem Left$(txtType.Text, 5) + "#" + CStr(c) _
        + " (" + Trim$(Mid$(txtType.Text, 6)) + ")"
Next i
ComputeCosts
Label1.Visible = (List1.ListCount > 0)
End Sub

Private Sub cmdBuild_Click()

'Scan Build Queue
Dim BuildCount As Integer
Dim BuildID As String
Dim strID As String
Dim i As Integer
For i = 1 To List1.ListCount
    strID = Left$(List1.List(i - 1), InStr(List1.List(i - 1), "#") - 1)
    If strID = BuildID Then
        BuildCount = BuildCount + 1
    ElseIf Len(BuildID) = 0 Then
        BuildID = strID
        BuildCount = 1
    Else
        BuildUnit BuildID, BuildCount
        BuildID = strID
        BuildCount = 1
    End If
Next i

'build the last one
BuildUnit BuildID, BuildCount

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If BuildType = "s" Then
    GetShips
ElseIf BuildType = "l" Then
    GetLandUnits
Else
    GetPlanes
End If

GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

List1.Clear
Me.Hide

End Sub
Private Sub BuildUnit(BuildID As String, BuildCount As Integer)
Dim strCmd As String

If BuildCount > 0 Then '180304 rjk: ensure the list is not empty, otherwise it generate a server problem.
    strCmd = "build " + BuildType + " " + strSector + _
                " " + BuildID + " " + CStr(BuildCount)
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If
End Sub

Private Sub cmdClear_Click()
List1.Clear
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Public Sub LoadTypebox(UnitClass As String)

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

If txtType.ListCount > 0 Then
    txtType.ListIndex = 0
End If
BuildType = UnitClass
ComputeCosts
End Sub


Private Sub Form_Activate()
Label1.Visible = (List1.ListCount > 0)
End Sub

Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
ComputeCosts
Label1.Visible = (List1.ListCount > 0)
End Sub


Public Sub ComputeCosts()

Dim i As Integer
Dim Cost As Long
Dim lcm As Long
Dim hcm As Long
Dim avail As Long
Dim Mil As Long
Dim gun As Long

For i = 0 To List1.ListCount - 1
    
    'get the build requirements
    rsBuild.Seek "=", BuildType, Trim$(Left$(List1.List(i), 5))
    If Not rsBuild.NoMatch Then
        'load costs
        Cost = Cost + rsBuild.Fields("cost")
        lcm = lcm + rsBuild.Fields("lcm")
        hcm = hcm + rsBuild.Fields("hcm")
        avail = avail + rsBuild.Fields("avail")
        Select Case BuildType
            Case "p"
                Mil = Mil + rsBuild.Fields("mil")
            Case "l"
                gun = gun + rsBuild.Fields("gun")
        End Select
    End If
Next i

If List1.ListCount > 0 Then   'make sure something was built
    'Fill cost string
    Frame2.Caption = CStr(List1.ListCount) + " Units in Queue"
    Label2(1).Caption = "Cost: $" + CStr(Cost)
    Label2(0).Caption = vbNullString
    'Fill resource expended string
    If lcm > 0 Then
        Label2(0).Caption = CStr(lcm) + " lcms, "
    End If
    If hcm > 0 Then
        Label2(0).Caption = Label2(0).Caption + CStr(hcm) + " hcms, "
    End If
    If avail > 0 Then
        Label2(0).Caption = Label2(0).Caption + CStr(avail) + " avail, "
    End If
    Select Case BuildType
        Case "p"
            If avail > 0 Then
                Label2(0).Caption = Label2(0).Caption + CStr(Mil) + " mil, "
            End If
        Case "l"
            If gun > 0 Then
                Label2(0).Caption = Label2(0).Caption + CStr(gun) + " guns, "
            End If
    End Select
    
    If Len(Label2(0).Caption) > 0 Then
        Label2(0).Caption = Left$(Label2(0).Caption, Len(Label2(0).Caption) - 2)
    End If
Else
    Frame2.Caption = "No Units in Queue"
    Label2(0) = vbNullString
    Label2(1) = vbNullString
End If

End Sub

