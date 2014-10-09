VERSION 5.00
Begin VB.Form frmPromptHelp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmPromptHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 08/11/03  efj Nothing in this file is used so I commented it all out

'Public HelpTopic As String
'
'Private Sub Form_Click()
'DisplayHelp
'End Sub
'
'Private Sub Form_Load()
'' Set parent for the toolbar to display on top of:
'Dim success As Long
'success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
'End Sub
'
'Private Sub Label1_Click()
'DisplayHelp
'End Sub
'
'Private Sub DisplayHelp()
'MsgBox HelpTopic
'End Sub


