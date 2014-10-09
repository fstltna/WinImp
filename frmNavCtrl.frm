VERSION 5.00
Begin VB.Form frmNavCtrl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Path"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Retreat Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      TabIndex        =   20
      Top             =   3840
      Width           =   2535
      Begin VB.CheckBox chkRetreat 
         Caption         =   "Set Retreat "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtRetreatPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "cond"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.ListBox lstReport 
      Height          =   3660
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9615
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1800
      TabIndex        =   8
      Top             =   3840
      Width           =   3495
      Begin VB.CommandButton cmdOption 
         Caption         =   "Halt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2400
         TabIndex        =   18
         Tag             =   "h"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Sweep"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2400
         TabIndex        =   17
         Tag             =   "m"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "List Ships"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Tag             =   "i"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Change Flagship"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   15
         Tag             =   "f"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Sonar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Tag             =   "s"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Tag             =   "v"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Bmap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   12
         Tag             =   "B"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   11
         Tag             =   "M"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Radar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Tag             =   "r"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "Look"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Tag             =   "l"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frameDir 
      Caption         =   "Movement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
      Begin VB.CommandButton cmdDir 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   1
         Top             =   1200
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmNavCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'060304 rjk: Added Sonar parsing to Navigation form.
'230406 rjk: Fixed the refresh of the units for 4.3.X servers.
'160706 rjk: Added AutoView for navigate and march.

Public Looking As Boolean
Public Sonar As Boolean

Public AutoLooking_when_notAutoStopping As Boolean
Public AutoViewing_when_notAutoStopping As Boolean
Public RefreshBmapOnStop As Boolean
Private fooLook As Boolean

Dim TextLines As Integer
Dim InputInhibited As Boolean

Private Sub cmdDir_Click(Index As Integer)
'The buttons caption field contains the code to send
SendCommand LCase(cmdDir(Index).Caption)
If cmdDir(Index).Caption <> "H" Then
    fooLook = True
Else
    fooLook = False
End If

End Sub

Public Sub callback()
'drk@unxsoft.com 7/13/05
If AutoLooking_when_notAutoStopping And fooLook Then
    Looking = True
    SendCommand "l"
End If
If AutoViewing_when_notAutoStopping And fooLook Then
    SendCommand "v"
End If
fooLook = False
End Sub

Private Sub cmdOption_Click(Index As Integer)

'If this is a look, set the looking flag so info will get written to
'enemy information database
If cmdOption(Index).Tag = "l" Then
    Looking = True
End If
If cmdOption(Index).Tag = "s" Then
    Sonar = True
End If

'The buttons tag field contains the code to send
SendCommand cmdOption(Index).Tag
End Sub

Private Sub SendCommand(strCmd As String)

'Send Command to Server
frmEmpCmd.SendStringtoServer strCmd
frmEmpCmd.List1.AddItem strCmd
AddReportLine strCmd

EnableButtons False

'if this is a movement, set the retreat path
If InStr("yugjbn", strCmd) > 0 Then
    txtRetreatPath.Text = ReversePath(strCmd) + txtRetreatPath.Text
    If Len(txtRetreatPath.Text) > 5 Then
        txtRetreatPath.Text = Left$(txtRetreatPath.Text, 5)
    End If
End If

'if we are done moving, submit the retreat command
If Left$(strCmd, 1) = "h" And chkRetreat.Value = vbChecked And Len(txtRetreatPath) > 0 And Len(frmDrawMap.LastRetreatUnits) > 0 Then
    strCmd = frmDrawMap.LastRetreatType + "retreat " + frmDrawMap.LastRetreatUnits + " " + txtRetreatPath.Text + " " + Left$(Combo1.Text, 1)
    frmEmpCmd.SubmitEmpireCommand strCmd, False
End If
If Left$(strCmd, 1) = "h" Then
    frmEmpCmd.SubmitEmpireCommand "db1", False
    frmEmpCmd.SubmitEmpireCommand "map *", False
    frmEmpCmd.SubmitEmpireCommand "bmap *", False
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If
End Sub

Private Sub Form_Activate()

'If mimimized for some reason, pop it up
If Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
    Exit Sub
End If

End Sub
Private Sub cmdOption_KeyPress(Index As Integer, KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub
Private Sub cmdDir_KeyPress(Index As Integer, KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
'if the form is closed while server is waiting for
'a response, send an "h" to halt movement
If frmEmpCmd.ServerQuery Then
    SendCommand "h"
End If

'drk 7/13/02: is there a better way to do this?
If RefreshBmapOnStop Then
    frmEmpCmd.SubmitEmpireCommand "db1", False
    frmEmpCmd.SubmitEmpireCommand "map *", False
    frmEmpCmd.SubmitEmpireCommand "bmap *", False
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If

End Sub

Private Sub lstReport_KeyPress(KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If InputInhibited Then
    Exit Sub
End If

'If a legit command is hit, process it
If InStr("yughjbnvirslMBfm", Chr(KeyAscii)) > 0 Then
    SendCommand Chr$(KeyAscii)
End If


End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2
TextLines = 0
EnableButtons

If Len(frmDrawMap.LastRetreatUnits) > 0 Then
    LoadRetreatBox Combo1
    Combo1.ListIndex = 0
    'If Len(frmDrawMap.LastRetreatString) > 0 Or frmDrawMap.mnuOptions(10).Checked = True Then
    'drk@unxsoft.com 10/25/02: the setting on the frmNavCtrl Box should not be overriden by the options menu.
    'The options menu should provide a default setting for that form.
    If frmDrawMap.DoRetreat Then
        chkRetreat.Value = 1
    Else
        chkRetreat.Value = 0
    End If
Else
    chkRetreat.Value = 0
    Frame1.Visible = False
End If

txtRetreatPath.Text = frmDrawMap.LastRetreatString
End Sub

'Public Sub ClearReport()   removed efj 8/2003
''clear text box
'lstReport.Clear
'' TextMaxWidth = 0
'TextLines = 0
'End Sub

Public Sub AddReportLine(strLine As String)
'Add line to text box
lstReport.AddItem strLine
TextLines = TextLines + 1

End Sub


Public Sub EnableButtons(Optional bState As Boolean)

Dim X As Integer

'Default is true
If IsMissing(bState) Then
    bState = True
End If

'First set direction buttons
For X = 0 To 6
    cmdDir(X).Enabled = bState
Next X

'Then set options
For X = 0 To 9
    cmdOption(X).Enabled = bState
Next X

InputInhibited = Not bState

End Sub

