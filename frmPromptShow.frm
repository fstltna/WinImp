VERSION 5.00
Begin VB.Form frmPromptShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
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
      Left            =   5640
      TabIndex        =   17
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtTechLevel 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      Height          =   1335
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option6 
         Caption         =   "statistics"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "capacities"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "build"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option10 
         Caption         =   "nuke"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "bridge tower"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "sector"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "bridge span"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ship"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "plane"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "land"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Show"
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
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Tech Level (Optional)"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "frmPromptShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 08/11/03  efj     Commented everything out, nothing in this file is used

'Private Sub cmdCancel_Click()
'Unload Me
'End Sub
'
'Private Sub cmdHelp_Click()
'frmDrawMap.DisplayPromptHelp Label2.Caption
'End Sub
'
'Public Sub cmdOk_Click()
'Dim strCmd As String
'Dim x As Integer
'On Error Resume Next
'
''Start with first list box results
'If Option1.Value Then
'    strCmd = "land"
'ElseIf Option2.Value Then
'    strCmd = "plane"
'ElseIf Option3.Value Then
'    strCmd = "ship"
'ElseIf Option7.Value Then
'    strCmd = "bridge"
'ElseIf Option9.Value Then
'    strCmd = "tower"
'ElseIf Option10.Value Then
'    strCmd = "nuke"
'Else
'    strCmd = "sector"
'End If
'
''Add result from second list box
'If Option7.Value Or Option9.Value Or Option4.Value Then
'    strCmd = strCmd + " build"
'ElseIf Option5.Value Then
'    strCmd = strCmd + " capacities"
'Else
'    strCmd = strCmd + " statistics"
'End If
'
'strCmd = "show " + strCmd
'
''Add tech level if input
'If Len(txtTechLevel.Text) > 0 Then
'    x = val(txtTechLevel.Text)
'    If x < 1 Then
'        MsgBox "Tech String not valid", vbExclamation + vbOKOnly, "Entry Error"
'        Exit Sub
'    End If
'    strCmd = strCmd + " " + CStr(val(txtTechLevel.Text))
'End If
'
'frmEmpCmd.SubmitEmpireCommand strCmd, True
'Unload Me
'End Sub
'
'Private Sub Form_Load()
''Make Stay always on top
'Dim success As Long
'success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'Set frmDrawMap.PromptForm = Nothing
'frmDrawMap.PromptUp = False
'End Sub
'
'Private Sub Option6_Click()
'Dim n As Integer
'If val(txtTechLevel.Text) < 1 And TechLevel > 0 Then
'    n = CInt(TechLevel)
'    If CSng(n) > TechLevel Then n = n - 1
'    txtTechLevel.Text = CStr(n)
'End If
'End Sub
