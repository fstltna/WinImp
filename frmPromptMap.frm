VERSION 5.00
Begin VB.Form frmPromptMap 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1590
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
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
      TabIndex        =   8
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Planes"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Tag             =   "p"
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Lands"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Tag             =   "l"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Ships"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Tag             =   "s"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Map"
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
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPromptMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092803 rjk: General Reformatting. Made Okay the default
'            Added Grid Unit selection on start up and field
'            selection.  Added reset to normal sector display upon
'            exit.  Added initial field selection.
'            Move field selection logic to form from frmDrawMap
'            Remove spaces between options so they were work

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim n As Integer
strCmd = LCase$(Trim$(Label2.Caption)) + " "
If txtOrigin.Visible Then
    strCmd = strCmd + txtOrigin.Text + " " ' 092403 rjk: Removed the space between options as they are not accepted
Else
    strCmd = strCmd + txtUnit.Text + " " ' 092403 rjk: Removed the space between options as they are not accepted
End If

'Check the check boxes for display options
For n = 0 To 2
    If Check1(n).Value = vbChecked Then
        strCmd = strCmd + Check1(n).Tag ' 092403 rjk: Removed the space between options as they are not accepted
    End If
Next n

frmEmpCmd.SubmitEmpireCommand strCmd, True

Unload Me
End Sub

Private Sub Form_Activate()
Select Case LCase$(Trim$(Label2.Caption))
Case "smap", "lmap", "pmap", "sbmap", "lbmap", "pbmap"
    frmPromptMap.txtUnit.Visible = True
    frmPromptMap.txtUnit.SetFocus
    frmPromptMap.txtOrigin.Visible = False
Case Else
    frmPromptMap.txtUnit.Visible = False
    frmPromptMap.txtOrigin.Visible = True
    frmPromptMap.txtOrigin.SetFocus
End Select
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
'Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
frmDrawMap.DisplayFirstSectorPanel
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

Private Sub txtOrigin_GotFocus()
frmDrawMap.DisplayFirstSectorPanel
End Sub

Private Sub txtUnit_GotFocus()
Select Case LCase$(Trim$(Label2.Caption))
Case "smap", "sbmap"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "lmap", "lbmap"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
Case "pmap", "pbmap"
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
End Select
End Sub
