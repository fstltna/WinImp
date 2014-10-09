VERSION 5.00
Begin VB.Form frmPromptConsider 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Consider"
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptConsider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String

strCmd = LCase(Label2.Caption)
strCmd = strCmd + " " + txtNum.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True

Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
Dim success As Long
success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

