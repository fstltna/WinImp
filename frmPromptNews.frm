VERSION 5.00
Begin VB.Form frmPromptNews 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOutput 
      Caption         =   "Report Output"
      Height          =   615
      Left            =   3600
      TabIndex        =   16
      Top             =   1920
      Width           =   3975
      Begin VB.OptionButton optNone 
         Caption         =   "None"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optHourly 
         Caption         =   "Hourly Activity"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optNewsPaper 
         Caption         =   "Newspaper"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
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
      Left            =   7320
      TabIndex        =   14
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3255
      Begin VB.CheckBox chkTelegramCount 
         Caption         =   "Analyse Telegram Activity (Show Relations Grid)"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cbVictim 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cbActor 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox chkVictim 
         Caption         =   "Victim"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chkActor 
         Caption         =   "Actor"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkHeadlines 
         Caption         =   "Show Headlines only"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "View News Entries"
      Height          =   1215
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   2655
      Begin VB.ComboBox cbDays 
         Height          =   315
         ItemData        =   "frmPromptNews.frx":0000
         Left            =   1320
         List            =   "frmPromptNews.frx":0002
         TabIndex        =   5
         Text            =   "cbDays"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "for the last"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Since Last Time I checked"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "day(s)"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Newspaper"
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
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmPromptNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: added Option Explicit and removed dead variables
'081203 efj: added missing variable definitions
'092703 rjk: general reformatting
'113003 rjk: Added Hourly Activity Report.  Added check box for starting relations grid.
'120203 rjk: Switched output selection to an option box to make clearer.

Private Sub cbDays_Change()
Option2.Value = True
End Sub

Private Sub cbDays_Click()
Option2.Value = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim strConnect As String

'make sure number of days are valid
If Option2.Value Then
    If Val(cbDays.Text) = 0 Then
        MsgBox "Number of days not valid", vbOKOnly + vbExclamation, "Entry Error"
        cbDays.SetFocus
        Exit Sub
    End If
End If

'make sure actor name/number
If chkActor.Value = 1 Then
    If Len(cbActor.Text) = 0 Then
        MsgBox "Must choose nation for Actor", vbOKOnly + vbExclamation, "Entry Error"
        cbVictim.SetFocus
        Exit Sub
    End If
End If
        
'make sure victim name/number
If chkVictim.Value = 1 Then
    If Len(cbVictim.Text) = 0 Then
        MsgBox "Must choose nation for Victim", vbOKOnly + vbExclamation, "Entry Error"
        cbVictim.SetFocus
        Exit Sub
    End If
End If
        
'112903 rjk: Popup the Relations Grid to show the Telegram results.
If chkTelegramCount <> vbUnchecked Then
    frmRelationsGrid.Show vbModeless, frmDrawMap
    frmRelationsGrid.go
End If

'112903 rjk: Added By Hour Activity Report
If optNewsPaper Then
    frmEmpCmd.SubmitEmpireCommand "np1 Newspaper", False
ElseIf optHourly Then
    frmEmpCmd.SubmitEmpireCommand "np1 Hourly", False
ElseIf optNone Then
    frmEmpCmd.SubmitEmpireCommand "np1 None", False
End If

'Check if newspaper or just headlines
If chkHeadlines = 1 Then
    strCmd = "headlines"
Else
    strCmd = "newspaper"
End If

'Check for number of days
If Option2 Then
    strCmd = strCmd + " " + CStr(Val(frmPromptNews.cbDays.Text))
End If

'Check for optional victim & actor prompts
If chkActor = 1 Then
    strCmd = strCmd + " ?actor=" + CStr(frmPromptNews.cbActor.ItemData(frmPromptNews.cbActor.ListIndex))
    strConnect = "&"
Else
    strConnect = " ?"
End If

If chkVictim = 1 Then
    strCmd = strCmd + strConnect + "victim=" + CStr(frmPromptNews.cbVictim.ItemData(frmPromptNews.cbVictim.ListIndex))
End If

'Submit the command
frmEmpCmd.SubmitEmpireCommand strCmd

'112903 rjk: Added By Hour Activity Report
If Not optNewsPaper Then
    frmEmpCmd.SubmitEmpireCommand "np2", False
End If

Unload Me
End Sub

Private Sub Form_Load()
Dim X As Integer    ' added 8/12/03 efj
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

For X = 1 To 14
    Me.cbDays.AddItem CStr(X)
Next X

cbDays.Text = CStr(1)
Option1.Value = True

'Fill list box with nation names
Nations.FillListBox cbActor
Nations.FillListBox cbVictim
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

