VERSION 5.00
Begin VB.Form frmPromptFollow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   7320
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
      TabIndex        =   7
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Leader"
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      Begin VB.TextBox txtUnit2 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Followers"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Follow"
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
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: unused file, commented everything out
'092303 rjk: Added back and added menu click event to call the form
'            added unit grid selection
'102203 rjk: Added Unload Me' to follow form to make consist with other forms when presenting
'            the Okay button.

Public Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String
Dim num As Integer
Dim x As Integer

strCmd = "follow " + txtUnit.Text + " " + txtUnit2.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
txtUnit.Visible = True
frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
frmDrawMap.DisplayFirstSectorPanel
End Sub

