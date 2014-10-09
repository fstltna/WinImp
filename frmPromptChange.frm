VERSION 5.00
Begin VB.Form frmPromptChange 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6120
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
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   1095
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "country name"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "password"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "New Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Change"
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
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmPromptChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Added Option Explicit and removed dead variables
'092803 rjk: general reformatting.
'100703 rjk: Nation list now in relations only so do not need "report *"
'100703 rjk: Does not accept a string longer than 19 characters for country name.
'100703 rjk: remove any space in the front and rear of the country name to make matching easier.
'180206 rjk: Replace relation with GetCountryList.

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

strCmd = "change "
If Option1 Then
    txtNew.Text = Trim$(txtNew.Text) '100703 rjk: remove any space in the front and rear to make matching easier.
    If Len(txtNew.Text) > 19 Then '100703 rjk: Does not accept a string longer than 19 characters.
        txtNew.Text = Left$(txtNew.Text, 19)
    End If
    frmEmpCmd.CountryName = txtNew.Text '100703 rjk: Updated the name so relation grid will work
    strCmd = strCmd + "country " + Chr$(34) + txtNew.Text + Chr$(34)
    frmEmpCmd.SubmitEmpireCommand strCmd, True
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetCountryList
    frmEmpCmd.SubmitEmpireCommand "db2", False
Else
    strCmd = strCmd + "representative " + Chr$(34) + txtNew.Text + Chr$(34)
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If
'database update
Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub
