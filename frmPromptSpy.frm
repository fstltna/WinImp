VERSION 5.00
Begin VB.Form frmPromptSpy 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1155
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "from"
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
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Spy"
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
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String

strCmd = "spy " + " " + txtOrigin.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
Dim sucess As Long
success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 
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


