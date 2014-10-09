VERSION 5.00
Begin VB.Form frmPromptIsland 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbNations 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDest 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Island Setup Tool"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "from"
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
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmPromptIsland"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092803 rjk: general reformatting.
'100603 rjk: Changed OK/Cancel labels for Set Owner to be less confusing.
'100603 rjk: Added initial field selection.

Private Done As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
' Dim strCmd As String       removed 3/2003 efj
Dim x1 As Integer
Dim X2 As Integer
Dim y1 As Integer
Dim Y2 As Integer

If Not ParseSectors(x1, y1, txtOrigin.Text) Then
    MsgBox "Must fill in 'from' sector", , "Error"
    txtOrigin.SetFocus
    Exit Sub
End If

If Not ParseSectors(X2, Y2, txtDest.Text) Then
    If Not Label2 = "Set Owner" Then
        MsgBox "Must fill in 'to' sector", , "Error"
        txtDest.SetFocus
        Exit Sub
    Else
        X2 = x1
        Y2 = y1
    End If
End If


If Label2 = "Island Develop Tool" Then
    IslandDevelop x1, y1, X2, Y2
ElseIf Label2 = "Island Setup Tool" Then
    BlitzSetup False, x1, y1, X2, Y2
ElseIf Label2 = "Set Owner" Then
    SetOwner x1, y1, X2, Y2, cbNations.ItemData(cbNations.ListIndex)
End If

If Not Label2 = "Set Owner" Then
    Unload Me
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next

If Done = 0 Then
    If Label2 = "Set Owner" Then '100603 rjk: Changed labels for Set Owner to be less confusing.
        cmdOK.Caption = "Set Owner"
        cmdCancel.Caption = "Exit"
        cbNations.SetFocus
    Else
        If Len(txtOrigin.Text) = 0 Or Len(txtDest.Text) > 0 Then
            txtOrigin.SetFocus
        Else
            txtDest.SetFocus
        End If
    End If
    Done = 1
End If

End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long       removed 3/2003 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

Done = 0 '100603 rjk: Ensure the form runs activated logic each startup.

'Fill list box with nation names
Nations.FillListBox cbNations
cbNations.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

On Error Resume Next
'If first box has been "clicked", more cursor to second box
If ParseSectors(sx, sy, txtOrigin.Text) Then
    txtDest.SetFocus
End If
End Sub
