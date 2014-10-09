VERSION 5.00
Begin VB.Form frmPromptOrigin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1560
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShiftOnly 
      Caption         =   "Shift Origin in the Database without submitting the origin command"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
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
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Change"
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
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "Origin"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "to"
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
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmPromptOrigin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Added Option Explicit and removed dead variables
'092703 rjk: Set initial field, general reformatting
'200404 rjk: Added the ability to shift database origin without submitting
'            a command to server.

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim sx As Integer
Dim sy As Integer

If chkShiftOnly.Value <> vbUnchecked Then
    If ParseSectors(sx, sy, txtOrigin.Text) Then
        ShiftOrigin sx, sy
        OriginX = OriginX - sx
        OriginY = OriginY - sy
        frmDrawMap.SelX = frmDrawMap.SelX - sx
        frmDrawMap.SelY = frmDrawMap.SelY - sy
        frmDrawMap.FillSectorBox frmDrawMap.SelX, frmDrawMap.SelY
        'OriginY must be an even number
        If (CInt(OriginY \ 2)) * 2 <> OriginY Then
            OriginY = OriginY + 1
            OriginX = OriginX + 1
        End If
        frmDrawMap.DrawHexes
    End If
Else
    strCmd = LCase$(Label2.Caption) + " " + txtOrigin.Text
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If
Unload Me
End Sub

Private Sub Form_Activate()
If frmEmpireLogin.bOffline Then
    chkShiftOnly.Value = vbGrayed
End If
txtOrigin.SetFocus
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
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
