VERSION 5.00
Begin VB.Form frmPromptOffer 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "announcements"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "mail"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDur 
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      Caption         =   "treaty"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "loan"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cbNations 
      Height          =   315
      Left            =   3000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Duration"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Rate"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Offer"
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
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmPromptOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092403 rjk: Switched to label1 to left justified and moved txtNum 500 to make room for number from Consider
'092703 rjk: General reformatting
'100903 rjk: Moved the field logic to the form.  Modified the form to 'reject accept' and 'reject reject' commands

Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
' Dim strCmd2 As String    8/2003 efj  removed

'Build Loan offer
strCmd = Trim$(LCase$(Label2.Caption))
If strCmd = "offer" Then
    If Option1.Value Then
        strCmd = strCmd + " loan " + cbNations.Text
        If txtAmount.Visible = True Then
            If Len(txtAmount) > 0 Then
                strCmd = strCmd + " " + txtAmount.Text
                If Len(txtDur) > 0 Then
                    strCmd = strCmd + " " + txtDur.Text
                    If Len(txtRate) > 0 Then
                        strCmd = strCmd + " " + txtRate.Text
                    End If
                End If
            End If
        Else
            strCmd = strCmd + " " + txtNum.Text
        End If
    Else
        strCmd = strCmd + " treaty " + cbNations.Text
    End If
ElseIf strCmd = "consider" Then
    If Option1.Value Then
        strCmd = strCmd + " loan " + txtNum.Text
    Else
        strCmd = strCmd + " treaty " + txtNum.Text
    End If
Else '100903 rjk: Added 'reject accept' and 'reject reject' commands
    strCmd = "reject " + strCmd
    If Option1.Value Then
        strCmd = strCmd + " loan"
    ElseIf Option2.Value Then
        strCmd = strCmd + " treaties"
    ElseIf Option3.Value Then
        strCmd = strCmd + " mail"
    ElseIf Option4.Value Then
        strCmd = strCmd + " announcements"
    End If
    If Len(cbNations.Text) > 0 Then
        strCmd = strCmd + " " + cbNations.Text
    End If
End If

frmEmpCmd.SubmitEmpireCommand strCmd, True

Unload Me
End Sub

Private Sub Form_Activate()
Select Case Trim$(LCase$(Label2.Caption))
Case "offer"
Case "consider"
    txtNum.Visible = True
    txtNum.Text = vbNullString
    '092403 rjk: Added 500 to make room for number
    txtNum.Move cbNations.Left + 500, cbNations.top
    cbNations.Visible = False
    Label1.Caption = "number"
    LoanPromptsVisible False
Case "reject", "accept" '100903 rjk: Added additional fields for reject commands
    LoanPromptsVisible False
    Option3.Visible = True
    Option4.Visible = True
End Select
 
If Trim$(LCase$(Label2.Caption)) <> "consider" Then
    'Fill list box with nation names
    Nations.FillListBox cbNations
    cbNations.ListIndex = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next 'In case nation list is empty
'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
 
'Put form in proper place
Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
top = frmDrawMap.top + frmDrawMap.Height - Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub Option1_Click()

If LCase$(Trim$(Label2.Caption)) = "offer" Then
    LoanPromptsVisible True
End If

End Sub

Private Sub Option2_Click()
    LoanPromptsVisible False
End Sub

'100903 rjk: Added to manage the LoanPrompts fields on the screen.
Private Sub LoanPromptsVisible(bVisible As Boolean)
Label3.Visible = bVisible
Label4.Visible = bVisible
Label5.Visible = bVisible
txtAmount.Visible = bVisible
txtRate.Visible = bVisible
txtDur.Visible = bVisible
End Sub
