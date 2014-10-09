VERSION 5.00
Begin VB.Form frmPromptLand 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5175
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
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2895
      Begin VB.TextBox txtParm 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "in"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Show"
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Land"
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
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptLand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Added Option Explicit
'092803 rjk: General reformatting.  Add initital field selection.

Public cmd As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

strCmd = LCase$(cmd) + " " + txtOrigin.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
Unload Me
End Sub

Private Sub Form_Activate()
txtOrigin.SetFocus
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
cmd = "land"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub
