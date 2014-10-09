VERSION 5.00
Begin VB.Form frmServerQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Query"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtRespond 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   7095
   End
   Begin VB.ListBox lstMsgs 
      Height          =   2460
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Response"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "frmServerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables and procedures
'093003 rjk: general reformatting.


Dim TextMaxWidth As Integer
Dim TextLines As Integer

Private Sub cmdAbort_Click()
txtRespond.Text = "ctld"
Me.Hide
End Sub

Private Sub cmdOK_Click()
If Len(txtRespond.Text) = 0 Then
    Exit Sub
End If

Me.Hide
End Sub

Private Sub Form_Activate()
Dim nLen As Long

If TextMaxWidth > 0 Then
    If TextMaxWidth > ((Screen.Width * 4) / 5) Then
        TextMaxWidth = (Screen.Width * 4) / 5
    End If
            
    Me.Width = TextMaxWidth * 1.15
    Me.Left = (Screen.Width - Me.Width) \ 2
End If

If TextLines > 1 Then
    nLen = Me.TextHeight("XX") * (TextLines + 2)
    If lstMsgs.Height > ((Screen.Height * 3) / 4) Then
        lstMsgs.Height = (Screen.Height * 3) / 4
    Else
        lstMsgs.Height = nLen
    End If
    
    Me.Height = lstMsgs.Height + (3 * cmdOK.Height)
    Me.top = (Screen.Height - Me.Height) \ 2
End If

frmServerQuery.SetFocus
txtRespond.SetFocus
lstMsgs.TopIndex = lstMsgs.ListCount - 1
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)   removed 08/12/03  efj
'
'If KeyAscii = 13 Then
'    cmdOK.Value = True
'End If
'
'End Sub

Private Sub Form_Resize()
' Dim x As Integer  removed 8/12/03 efj
Dim n As Integer

If WindowState <> vbNormal Then
    Exit Sub
End If

' x = Me.TextHeight("XX")  removed 8/12/03 efj
n = lstMsgs.ListCount       ' changed from m to n  08/11/03 efj
If n > 10 Then
    n = 10
End If

If Me.Width < 3 * cmdOK.Width Then
    Me.Width = 3 * cmdOK.Width
End If

If Me.Height < 5 * cmdOK.Height Then
    Me.Height = 5 * cmdOK.Height
End If

txtRespond.Width = lstMsgs.Width


Me.Label1.top = Me.lstMsgs.top + lstMsgs.Height + Label1.Height

Me.txtRespond.top = Label1.top + 2 * Label1.Height

cmdOK.Move (Me.Width - 2.5 * cmdOK.Width) \ 2, txtRespond.top + txtRespond.Height + cmdOK.Height * 0.5
cmdAbort.Move cmdOK.Left + 1.5 * cmdOK.Width, cmdOK.top

Me.Height = cmdOK.top + 3 * cmdOK.Height

If Me.Height < 5 * cmdOK.Height Then
    Me.Height = 5 * cmdOK.Height
End If

Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2
End Sub

