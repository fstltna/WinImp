VERSION 5.00
Begin VB.Form frmPromptWipe 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1260
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4560
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
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "all thesholds from"
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
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Wipe"
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
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "frmPromptWipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'270903 rjk: General reformatting
'171003 rjk: Added Strength fields to Sector database
'201103 rjk: Added option to control strength updates
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

strCmd = "wipe " + " " + txtOrigin.Text
frmEmpCmd.SubmitEmpireCommand strCmd, True
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

'092703 rjk: Added to draw focus to Origin on startup
Private Sub Form_Activate()
txtOrigin.SetFocus
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long    removed 8/2003 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
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
