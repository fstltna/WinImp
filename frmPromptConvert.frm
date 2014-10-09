VERSION 5.00
Begin VB.Form frmPromptConvert 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5655
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
      TabIndex        =   7
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkReserve 
      Caption         =   "Put on Reserves"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Convert"
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
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "civilians in"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: added Option Explcit and removed dead variables
'092803 rjk: Fill in txtNum based on the origin and command.
'            Added initial field selection. General reformatting.
'093003 rjk: Blank the txtNum field if it is not a valid origin.
'            Only move focus to txtNum field if origin sector is valid.
'            Update the mil reserve level after enlist and demobilize commands.
'093003 rjk: removed the LCase for command determination and adjust the corresponding text.
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace nation with GetNations.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String

If Label2.Caption = "Realm" Then '093003 rjk: removed the LCase and adjust the text.
    strCmd = LCase$(Label2.Caption) + " " + txtNum.Text + " " + txtOrigin.Text
    frmEmpCmd.SubmitEmpireCommand strCmd, True
Else
    strCmd = LCase$(Label2.Caption) + " " + txtOrigin.Text + " " + txtNum.Text
    'Handle reserve check box for demobilization
    If chkReserve.Visible Then
        If chkReserve.Value = vbChecked Then
            strCmd = strCmd + " y"
        Else
            strCmd = strCmd + " n"
        End If
    End If
        
    frmEmpCmd.SubmitEmpireCommand strCmd, True
    'database update
    frmEmpCmd.SubmitEmpireCommand "db1", False
    '093003 rjk: Update the mil reserve level after enlist or demobilize commands
    If Label2.Caption = "Enlist" Or Label2.Caption = "Demobilize" Then
        GetNation
    End If
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If
Unload Me
End Sub

Private Sub Form_Activate()
txtNum.SetFocus
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

Private Sub txtOrigin_Change()
Dim secx As Integer
Dim secy As Integer

If ParseSectors(secx, secy, txtOrigin.Text) Then
    rsSectors.Seek "=", secx, secy
    If Not rsSectors.NoMatch Then
        Select Case Label2.Caption
        Case "Grind"
            txtNum = CStr(rsSectors.Fields("bar"))
        Case "Realm", "Enlist"
            txtNum = vbNullString
        Case "Convert"
            txtNum = CStr(rsSectors.Fields("civ"))
        Case "Demobilize"
            txtNum = CStr(rsSectors.Fields("mil"))
        End Select
        If txtNum.Visible Then '093003 rjk: only move focus if sector is valid
            txtNum.SetFocus
        End If
    Else
        txtNum = vbNullString
    End If
Else '093003 rjk: Blank the field if it is not a valid origin.
    txtNum = vbNullString
End If
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
