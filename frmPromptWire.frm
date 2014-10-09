VERSION 5.00
Begin VB.Form frmPromptWire 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
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
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Clear Wire Reports"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "View Wire Entries"
      Height          =   1455
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3015
      Begin VB.ComboBox cbDays 
         Height          =   315
         ItemData        =   "frmPromptWire.frx":0000
         Left            =   1560
         List            =   "frmPromptWire.frx":0002
         TabIndex        =   5
         Text            =   "cbDays"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "for the last"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Since Last Time I checked"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "day(s)"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Wire"
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
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmPromptWire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: added "option explicit"  and definitions for undefined variables
'081103 efj: removed dead variables
'093003 rjk: general reformatting.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.

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
' Dim strConnect As String   removed efj 8/2003

'make sure number of days are valid
If Option2.Value Then
    If Val(cbDays.Text) = 0 Then
        MsgBox "Number of days not valid", vbOKOnly + vbExclamation, "Entry Error"
        cbDays.SetFocus
        Exit Sub
    End If
End If

strCmd = "wire"

'Check for number of days
If Option2 Then
    strCmd = strCmd + " " + CStr(Val(frmPromptWire.cbDays.Text))
End If

'Check if clearimg
If Check1.Value = vbChecked Then
    strCmd = strCmd + " y"
Else
    strCmd = strCmd + " n"
End If

'Submit the command
frmEmpCmd.SubmitEmpireCommand strCmd

Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long   removed efj 8/2003
' success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)      replaced with the following line to remove success 8/2003
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

Dim X As Integer    ' added definition efj 8/2003
For X = 1 To 14
    Me.cbDays.AddItem CStr(X)
Next X

cbDays.Text = CStr(1)
Option1.Value = True

'Set Check box
Check1.Value = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

