VERSION 5.00
Begin VB.Form frmPromtAttack 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Support"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "planes"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "artillery"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ships"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "forts"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Attack"
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
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPromtAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String
Dim strSupport As String
Dim x As Integer

strSupport = 0

For x = 0 To 3
    If Check1(x).Value = 1 Then
        strSupport = strSupport + " y"
    Else
        strSupport = strSupport + "n"
    End If
Next x

strCmd = "attack " + txtOrigin.Text

If strSupport <> " y y y y" Then
    strCmd = strCmd + strSupport
End If

frmEmpCmd.SubmitEmpireCommand strCmd, False
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


