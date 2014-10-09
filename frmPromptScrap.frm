VERSION 5.00
Begin VB.Form frmPromptScrap 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1845
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   615
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
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   1455
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "land"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ship"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "plane"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Units"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Price"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Scrap"
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
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmPromptScrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form or selecting fields
'            and when disactivating return to standard Sector display.
'            General reformatting. Added setting the initial field on start up.
'101403 rjk: added dump to pickup additional materials from scrap process
'101603 rjk: Added set to commands using this form
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace lost with GetLost.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Public strCmd As String

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
'Dim strCmd2 As String

If Option1.Value Then
    strCmd = strCmd + " ship "
'    strCmd2 = "sdump *"
ElseIf Option2.Value Then
    strCmd = strCmd + " land "
'    strCmd2 = "ldump *"
Else
    strCmd = strCmd + " plane "
'    strCmd2 = "pdump *"
End If

strCmd = strCmd + txtUnit.Text

If Left$(strCmd, 3) = "set" And Len(txtPrice.Text) > 0 And Len(txtUnit.Text) > 0 Then
    strCmd = strCmd + " " + txtPrice.Text
End If

frmEmpCmd.SubmitEmpireCommand strCmd, True

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'101403 rjk: added dump to pickup additional materials from scrap process
If InStr(strCmd, "scrap") > 0 Then
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
End If
'frmEmpCmd.SubmitEmpireCommand strCmd2, False
GetLost
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long    8/12/03 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
End Sub

Private Sub Option2_Click()
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
End Sub

Private Sub Option3_Click()
If txtUnit.Visible Then
    txtUnit.SetFocus
End If
End Sub

Private Sub txtUnit_GotFocus()
If Option1 Then
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
ElseIf Option2 Then
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
ElseIf Option3 Then
    frmDrawMap.SetUnitDisplay udPLANE, TYPE_ALL, False, False, False
End If
End Sub
