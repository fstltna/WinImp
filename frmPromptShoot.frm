VERSION 5.00
Begin VB.Form frmPromptShoot 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1725
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5640
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
      Left            =   5280
      TabIndex        =   9
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "civilians"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "un. workers"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "Shoot"
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
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmPromptShoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Added Option Explicit and removed dead variables
'092703 rjk: Set the quantity field based on the option and
'            origin.  General Reformatting.
'093003 rjk: only switch to txtNum field if origin is valid.
'093003 rjk: do not set txtNum for multiple sector locations.
'093003 rjk: Set the options based on the sector location.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'120203 rjk: Changed to make civ the default if present.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

If Option1 Then
    strCmd = "u "
Else
    strCmd = "c "
End If

strCmd = "shoot " + strCmd + txtOrigin.Text + " " + txtNum.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'frmEmpCmd.SubmitEmpireCommand "dump " + txtOrigin.Text, False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

Private Sub Form_Activate()
SetTextNum
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
SetTextNum
End Sub

Private Sub Option2_Click()
SetTextNum
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

Private Sub txtOrigin_Change()
'093003 rjk: Set the options based on the sector location.
Dim secx As Integer
Dim secy As Integer

If Len(txtOrigin.Text) > 0 Then
    If ParseSectors(secx, secy, txtOrigin.Text) Then
        rsSectors.Seek "=", secx, secy
        If Not rsSectors.NoMatch Then
            '120203 rjk: Switched the order to make civ the default if present.
            If CStr(rsSectors.Fields("uw")) > 0 Then
                Option1.Enabled = True
                Option1.Value = True
            Else
                Option1.Enabled = False
            End If
            If CStr(rsSectors.Fields("civ")) > 0 Then
                Option2.Enabled = True
                Option2.Value = True
            Else
                Option2.Enabled = False
            End If
        Else
            Option1.Enabled = False
            Option2.Enabled = False
        End If
    Else
        Option1.Enabled = True
        Option2.Enabled = True
        txtNum = vbNullString
    End If
Else
    Option1.Enabled = False
    Option2.Enabled = False
End If

SetTextNum
End Sub

Private Sub SetTextNum()
Dim secx As Integer
Dim secy As Integer

If Len(txtOrigin.Text) > 0 Then
    If ParseSectors(secx, secy, txtOrigin.Text) Then
        rsSectors.Seek "=", secx, secy
        If Not rsSectors.NoMatch Then
            If Option2 Then
                txtNum = CStr(rsSectors.Fields("civ"))
            ElseIf Option1 Then
                txtNum = CStr(rsSectors.Fields("uw"))
            Else
                txtNum = vbNullString
            End If
            If txtNum.Visible Then
                txtNum.SetFocus
            End If
        Else
            txtNum = vbNullString
            If txtOrigin.Visible Then '093003 rjk: only switch to txtNum field if origin is valid.
                txtOrigin.SetFocus
            End If
        End If
    'Else multiple sector selection, do not erase txtNum
    End If
Else
    txtNum = vbNullString
    If txtOrigin.Visible Then '093003 rjk: only switch to txtNum field if origin is valid.
        txtOrigin.SetFocus
    End If
End If
End Sub

