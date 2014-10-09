VERSION 5.00
Begin VB.Form frmPromptNation 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5415
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
      Left            =   5040
      TabIndex        =   9
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cbRelations 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox cbNations 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "des"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Relations"
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
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "with"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptNation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'120803 efj: removed dead variables
'270903 rjk: general reformatting
'011103 rjk: Form cleanup
'251103 rjk: Added relations update after doing a declare command.
'180206 rjk: Replace relation with GetCountryList.

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
' Dim strCmd2 As String    8/2003 efj  removed

strCmd = Trim$(LCase$(Label2.Caption))
If strCmd = "relations" Or strCmd = "accept" Then
    frmEmpCmd.SubmitEmpireCommand strCmd + " " + cbNations.Text, True
ElseIf strCmd = "declare" Then
    frmEmpCmd.SubmitEmpireCommand strCmd + " " + cbRelations.Text _
           + " " + cbNations.Text, True
ElseIf strCmd = "sharebmap" Then
    frmEmpCmd.SubmitEmpireCommand strCmd + " " + cbNations.Text + " " + txtOrigin.Text + " " + txtNum.Text, True
Else
    Exit Sub
End If
If strCmd = "declare" Then '112503 rjk: Added requesting a relations update to get the new declare status.
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetCountryList
    frmEmpCmd.SubmitEmpireCommand "db2", False
End If
Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 
'Fill list box with nation names
Nations.FillListBox cbNations

cbRelations.Clear
cbRelations.AddItem "alliance"
cbRelations.AddItem "friendly"
cbRelations.AddItem "neutrality"
cbRelations.AddItem "hostility"
cbRelations.AddItem "war"
cbRelations.ListIndex = 3   'default to hostile
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
