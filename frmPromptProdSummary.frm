VERSION 5.00
Begin VB.Form frmPromptProdSummary 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Show All Distribution Centers"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   4695
      Begin VB.CommandButton cmdAll 
         Caption         =   "Show All"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show Single Distribution Center Only"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4695
      Begin VB.CommandButton cmdOK 
         Caption         =   "Show Single"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dist. Center"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frameShow 
      Caption         =   "Problem Reports"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
      Begin VB.CheckBox chkBuild 
         Caption         =   "Build problems"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkIdle 
         Caption         =   "Idle Civs"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox ckOvrPop 
         Caption         =   "Overpopulation"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ckfood 
         Caption         =   "Food Shortage"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Production Summary"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPromptProdSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'092703 rjk: General reformatting
'102503 rjk: Replace Done with bFirstLoad, set txtOrigin as initial field
'            Removed DontMarkReportSectors, Enabled the user selected of reports

'Dim DontMarkReportSectors As Boolean 102503 rjk: Removed, always enabled
Dim bFirstLoad As Boolean

Private Sub cmdAll_Click()
    ProdSummaryReport 0, 1, Me.ckfood, Me.ckOvrPop, Me.chkIdle, Me.chkBuild
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'Dim strCmd As String    8/2003 efj  removed
Dim x1 As Integer
' Dim X2 As Integer    8/2003 efj  removed
Dim y1 As Integer
' Dim Y2 As Integer    8/2003 efj  removed

If Not ParseSectors(x1, y1, txtOrigin.Text) Then
    MsgBox "Must fill in 'from' sector", , "Error"
    txtOrigin.SetFocus
    Exit Sub
End If

ProdSummaryReport x1, y1, Me.ckfood, Me.ckOvrPop, Me.chkIdle, Me.chkBuild

Unload Me
End Sub

Private Sub Form_Activate()
If bFirstLoad Then
    txtOrigin.SetFocus
    bFirstLoad = False
End If
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

'102503 rjk: removed, frame always present
'Set Check Boxes
'If DontMarkReportSectors Then
'    Me.ckfood = 0
'    Me.ckOvrPop = 0
'    Me.chkIdle = 0
'    Me.chkBuild = 0
'Else
'    Me.ckfood = vbChecked
'    Me.ckOvrPop = vbChecked
'    Me.chkIdle = vbChecked
'    Me.chkBuild = vbChecked
'End If

bFirstLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

'102503 rjk: removed, frame always present
'Private Sub frameShow_Click()
'Toggle support
'DontMarkReportSectors = Not DontMarkReportSectors

'If DontMarkReportSectors Then
'    Me.ckfood = 0
'    Me.ckOvrPop = 0
'    Me.chkIdle = 0
'    Me.chkBuild = 0
'Else
'    Me.ckfood = vbChecked
'    Me.ckOvrPop = vbChecked
'    Me.chkIdle = vbChecked
'    Me.chkBuild = vbChecked
'End If
'End Sub
