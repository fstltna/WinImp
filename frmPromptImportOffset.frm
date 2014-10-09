VERSION 5.00
Begin VB.Form frmPromptImportOffset 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Offset"
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPromptImportOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'181006 rjk: Created for setting import offset for intelligence

Public strTelegram As String

Private Sub cmdCancel_Click()
strTelegram = vbNullString
Unload Me
End Sub

Private Sub cmdImport_Click()
Dim iOffsetX As Integer, iOffsetY As Integer
If ParseSectors(iOffsetX, iOffsetY, txtOrigin.Text) Then
    If Len(strTelegram) > 0 Then
        ImportIntelligenceTelegram strTelegram, iOffsetX, iOffsetY
        strTelegram = vbNullString
    Else
        ImportIntelligence iOffsetX, iOffsetY
    End If
    Unload Me
    frmDrawMap.DrawHexes
Else
    MsgBox "Invalid Sector", vbOKOnly
End If
End Sub

Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

