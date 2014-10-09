VERSION 5.00
Begin VB.Form frmPromptPath 
   Caption         =   "Select Path"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdConditional 
      Caption         =   "Conditional"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEndSector 
      Caption         =   "lblEndSector"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label7 
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
      Left            =   1680
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Path Cost"
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
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Ending Sector"
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
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblSector 
      Caption         =   "lblSector"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Total Path Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "path"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Sector"
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmPromptPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'092703 rjk: General reformatting
'101903 rjk: Added Condition Button for Multiple Sector Selection
'111403 rjk: Fixed so the keyboard can be used to select direction
'            Added Initial Field selection and focus changes on the return/home/reset buttons.

Public HoldSource As Form
' Public strSectors As String    8/2003 efj  removed

' Private ActivateUnitBox As Boolean    8/2003 efj  removed
Private Receiver As Control

Private Sub cmdCancel_Click()
frmDrawMap.FinishPath
Unload Me
End Sub

'101903 rjk: Return all the sectors that are conditional selected
Private Sub cmdConditional_Click()
frmDrawMap.FinishPath
Receiver.Text = GetConditionalSectors
Unload Me
End Sub

Private Sub cmdHome_Click()
Dim x1 As Integer
Dim X2 As Integer
Dim y1 As Integer
Dim Y2 As Integer
Dim i As Integer
Dim strPath As String

'Compute the path home, and add to path string
If ParseSectors(x1, y1, lblEndSector.Caption) Then
    If ParseSectors(X2, Y2, lblSector.Caption) Then
        strPath = EmpirePathDirection(X2 - x1, Y2 - y1)
        For i = 1 To Len(strPath)
            frmDrawMap.ExtendPath Mid$(strPath, i, 1)
        Next i
    End If
End If
cmdOK.SetFocus '111403 rjk: Draw the focus the Okay button to finish form.
End Sub

Private Sub cmdOK_Click()
frmDrawMap.FinishPath
Receiver.Text = txtPath.Text + "h"
Unload Me
End Sub

Private Sub cmdReset_Click()
frmDrawMap.FinishPath
frmDrawMap.StartPath lblSector.Caption
lblEndSector.Caption = lblSector.Caption
txtPath.Text = vbNullString
txtPath.SetFocus '111403 rjk: Draw the focus the txtPath so the keyboard works
End Sub

Private Sub cmdReturn_Click()
Dim i As Integer
Dim strPath As String
Dim ch As String

'Compute the path home, and add to path string
strPath = txtPath.Text
For i = Len(strPath) To 1 Step -1
    Select Case Mid$(strPath, i, 1)
        Case "y"
            ch = "n"
        Case "u"
            ch = "b"
        Case "j"
            ch = "g"
        Case "n"
            ch = "y"
        Case "b"
            ch = "u"
        Case "g"
            ch = "j"
        Case Else
            ch = vbNullString
    End Select
    frmDrawMap.ExtendPath ch
Next i
cmdOK.SetFocus '111403 rjk: Draw the focus the Okay button to finish form.
End Sub

Private Sub Form_Activate()
HoldSource.Enabled = False
If Not frmDrawMap.DrawingPath Then
    frmDrawMap.StartPath lblSector.Caption
End If
'101903 rjk: Only enable the button if the path is empty, it is multiple sector field, and conditional is active
If Len(txtPath.Text) = 0 And Receiver.Name = "txtMultPath" And ColorScheme = COLORSCHEME_CONDITIONAL Then
    cmdConditional.Visible = True
Else
    cmdConditional.Visible = False
End If
txtPath.SetFocus '111403 rjk: Draw the focus the txtPath so the keyboard works
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

'Set to recieve focus for clicks on DrawMap
Set HoldSource = frmDrawMap.PromptForm
Set frmDrawMap.PromptForm = Me

Set Receiver = HoldSource.ActiveControl
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = HoldSource
HoldSource.Enabled = True
On Error Resume Next
HoldSource.SetFocus

Set HoldSource = Nothing
End Sub

Private Sub txtPath_Change()
'Update path lenght
Label4 = CStr(Len(txtPath.Text))

'update the path cost
Dim c As Single
c = PathCost(Me.lblSector.Caption, txtPath.Text, MT_COMMODITY)
If c > 0 Then
    Label6.Caption = "Path Cost"
    Label7.Caption = CStr(CSng(CLng(c * 1000) / 1000))
Else
    Label6.Caption = vbNullString
    Label7.Caption = vbNullString
End If
'101903 rjk: Only enable the button if the path is empty, it is multiple sector field, and conditional is active
If Len(txtPath.Text) = 0 And Receiver.Name = "txtMultPath" And ColorScheme = COLORSCHEME_CONDITIONAL Then
    cmdConditional.Visible = True
Else
    cmdConditional.Visible = False
End If
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
Dim ch As String
Dim x As Integer

'If it is a valid direction entry, then add to path
ch = Chr$(KeyAscii)
x = InStr("gyujbn", ch) '111403 rjk: Reverse order to make the check work.
If x > 0 Then
    frmDrawMap.ExtendPath ch
    KeyAscii = 0
    Exit Sub
End If

'If its a backspace, then remove the last path entry
If KeyAscii = vbKeyBack Then
    frmDrawMap.RemoveFromPath
    KeyAscii = 0
    Exit Sub
End If

KeyAscii = 0
End Sub
