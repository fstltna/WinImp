VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chat (EmpireHub)"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.ComboBox cbAllies 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkBeep 
      Caption         =   "Beep"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.CheckBox chkTop 
      Caption         =   "Always on Top"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton optUser 
      Caption         =   "User"
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.OptionButton optAll 
      Caption         =   "All"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtMessage 
      Height          =   495
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2280
      Width           =   4965
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.Menu mnuBaseMenu 
      Caption         =   "Base Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyReportWindow 
         Caption         =   "Copy to Report Window"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'051904 rjk: Created
'052204 rjk: Modified to support Flash Chat as well.
'060604 rjk: Fixed the crash when window was made too small, also fixed the scale for the User combo box.
'130604 rjk: Changed the initial state for Flash window to country instead all.
'            Added MsgBox when no Country is selected.
'190804 rjk: Added reseting the token state for frmEmpCmd when closing the Chat window
'180206 rjk: Fix the check for empty country list
'090607 rjk: Fixed the All/Country selection to not change after it has been selected.
'            Changed the default for Chat.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

Private cChatColors(0 To 5) As Long
Private strUsers() As String

Private Sub cbAllies_Change()
optUser.Value = True
End Sub

Private Sub cbAllies_Click()
optUser.Value = True
End Sub

Private Sub chkTop_Click()
If chkTop.Value <> vbUnchecked Then
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
Else
    Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End If
End Sub

Private Sub Form_Activate()
If optUser.Value = False And optAll.Value = False Then
    If InStr(Me.Caption, "Flash") > 0 Then
        optUser.Value = True
    Else
        optAll.Value = True
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtMessage.Text) > 0 Then
        If InStr(Me.Caption, "Flash") > 0 Then
            If optAll Then
                frmEmpCmd.SubmitEmpireCommand "wall " + txtMessage.Text, False
                AddText "(" + frmEmpCmd.LoginUser + ")" + txtMessage.Text
                txtMessage = vbNullString
            ElseIf cbAllies.ListIndex <> -1 Then
                frmEmpCmd.SubmitEmpireCommand "flash " + cbAllies.List(cbAllies.ListIndex) + " " + txtMessage.Text, False
                AddText "(" + frmEmpCmd.LoginUser + ")" + txtMessage.Text
                txtMessage = vbNullString
            Else
                MsgBox "No Country Specified"
            End If
        Else
            If optAll Then
                frmEmpCmd.SubmitEmpireCommand "lwall " + txtMessage.Text, False
                txtMessage = vbNullString
            Else
                If Len(txtUser) > 0 Then
                    frmEmpCmd.SubmitEmpireCommand "lflash " + txtUser + " " + txtMessage.Text, False
                    AddText "(" + frmEmpCmd.LoginUser + ")" + txtMessage.Text
                    txtMessage = vbNullString
                Else
                    MsgBox "No User Specified"
                End If
            End If
        End If
    End If
    KeyAscii = 0
End If
End Sub

Public Sub AddText(strText As String)
Dim iPosStart As Integer
Dim iPosEnd As Integer
Dim strUser As String
Dim iColorIndex As Integer
Dim strMessage As String
Dim bFlash As Boolean

strMessage = strText

rtbText.SelStart = Len(rtbText.Text)
iPosStart = InStr(strMessage, "(")
iPosEnd = InStr(strMessage, ")")
If iPosEnd > iPosStart And iPosEnd > 0 Then
    If Mid$(strMessage, iPosStart, 2) = "(#" Then
        strUser = Mid$(strMessage, iPosStart + 2, iPosEnd - iPosStart - 2) 'country number
        bFlash = True
    Else
        strUser = Mid$(strMessage, iPosStart + 1, iPosEnd - iPosStart - 1) 'user name
        bFlash = False
    End If
    If Mid$(strMessage, iPosEnd, 2) = "):" Then
        strMessage = Right$(strMessage, Len(strMessage) - iPosEnd - 2)
    Else
        strMessage = Right$(strMessage, Len(strMessage) - iPosEnd)
    End If
    If Mid$(strMessage, 6, 1) = ":" And Mid$(strMessage, 9, 1) = ":" Then
        strMessage = LTrim$(Right$(strMessage, Len(strMessage) - 9))
    End If
    rtbText.SelColor = GetColor(strUser, bFlash)
    If Len(rtbText.Text) = 0 Then
        rtbText.SelText = strMessage
    Else
        rtbText.SelText = vbLf + strMessage
    End If
    If chkBeep <> vbUnchecked And frmOptions.bSound Then
        Beep
    End If
Else
    rtbText.Text = rtbText.Text + vbLf + "Incorrectly formed message"
End If
rtbText.SelStart = Len(rtbText.Text)
End Sub

Private Function GetColor(strUser As String, bFlash As Boolean) As Long
Dim iIndex As Integer

If bFlash And frmOptions.bFlashPlayerColors Then
    GetColor = PlayerColors(Val(strUser) Mod NUMBER_OF_PLAYER_COLORS)
    Exit Function
End If

For iIndex = 0 To UBound(strUsers)
    If strUsers(iIndex) = strUser Then
        GetColor = cChatColors(iIndex Mod (UBound(cChatColors) + 1))
        Exit Function
    End If
Next iIndex
iIndex = UBound(strUsers)
iIndex = iIndex + 1

ReDim Preserve strUsers(0 To iIndex)

strUsers(iIndex) = strUser
GetColor = cChatColors(iIndex Mod (UBound(cChatColors) + 1))
End Function


Private Sub Form_Load()
ReDim strUsers(0)
strUsers(0) = frmEmpCmd.LoginUser
cChatColors(0) = vbBlack
cChatColors(1) = vbBlue
cChatColors(2) = vbMediumGreen
cChatColors(3) = vbRed
cChatColors(4) = vbDarkBlue
cChatColors(5) = vbDarkGreen
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    Exit Sub
End If

If (Me.Height - txtMessage.Height - cbAllies.Height - txtMessage.Height - 850) < 0 Then
    Me.Height = txtMessage.Height + cbAllies.Height + txtMessage.Height + 850
End If
rtbText.Move 120, 120, Me.Width - 360, Me.Height - txtMessage.Height - cbAllies.Height - 850
txtMessage.Move rtbText.Left, rtbText.top + rtbText.Height + 120, rtbText.Width, txtMessage.Height
txtUser.Move txtUser.Left, Me.Height - 850
cbAllies.Move cbAllies.Left, Me.Height - 850
chkBeep.Move chkBeep.Left, Me.Height - 850
chkTop.Move chkTop.Left, Me.Height - 850
optAll.Move optAll.Left, Me.Height - 850
optUser.Move optUser.Left, Me.Height - 850
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Only unload if form Draw says OK
If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
    If (Not frmDrawMap.ShutDown) And Not (InStr(Me.Caption, "Flash") > 0 And Not frmOptions.bFlashChat) Then
        Cancel = 1
        Me.WindowState = vbMinimized
        If InStr(Me.Caption, "Flash") = 0 Then 'Reset if command gets locked but missing BTU response
            If frmEmpCmd.tokenStatus <> tsIdle Then
                frmEmpCmd.tokenStatus = tsIdle
            End If
            If frmEmpCmd.CmdinProgress = True Then
                frmEmpCmd.NextCommand
            End If
        End If
        Exit Sub
    End If
End If
End Sub

Public Sub ControlFlashForm()
Dim iFriendlies As Integer
Dim rAlly As Integer
Dim iAllies As Integer
Dim strFriendly() As String
Dim cnFriendly() As Integer
Dim iIndex As Integer

If frmOptions.bFlashChat And Not frmEmpireLogin.bOffline And Not frmDrawMap.ShutDown Then
    cnFriendly = Nations.GetCountryNumbers()
    If IsArrayEmpty(cnFriendly) < 0 Then
        Exit Sub
    End If
    For iIndex = LBound(cnFriendly) To UBound(cnFriendly)
        If cnFriendly(iIndex) <> CountryNumber Then
            rAlly = Nations.Relation(CountryNumber, cnFriendly(iIndex))
            If rAlly = iREL_ALLIED Or rAlly = iREL_FRIENDLY Then
                ReDim Preserve strFriendly(iFriendlies)
                strFriendly(iFriendlies) = Nations.NationName(cnFriendly(iIndex))
                iFriendlies = iFriendlies + 1
            End If
            If rAlly = iREL_ALLIED Then
                iAllies = iAllies + 1
            End If
        End If
    Next
    If iAllies = 0 And iFriendlies = 0 Then
        Exit Sub
    End If
    If frmEmpCmd.frmFlash Is Nothing Then
        Set frmEmpCmd.frmFlash = New frmChat
        frmEmpCmd.frmFlash.Caption = "Flash Chat"
        frmEmpCmd.frmFlash.cbAllies.Visible = True
        frmEmpCmd.frmFlash.optUser.Caption = "Country"
        frmEmpCmd.frmFlash.txtUser.Visible = False
        frmEmpCmd.frmFlash.Show vbModeless
    End If
    frmEmpCmd.frmFlash.cbAllies.Clear
    For iIndex = 0 To UBound(strFriendly)
        frmEmpCmd.frmFlash.cbAllies.AddItem strFriendly(iIndex)
    Next iIndex
    frmEmpCmd.frmFlash.cbAllies.ListIndex = -1
    If iAllies = 0 Then
        frmEmpCmd.frmFlash.optAll.Enabled = False
        frmEmpCmd.frmFlash.optUser.Value = True
    Else
        frmEmpCmd.frmFlash.optAll.Enabled = True
    End If
Else
    If Not (frmEmpCmd.frmFlash Is Nothing) Then
        Unload frmEmpCmd.frmFlash
        Set frmEmpCmd.frmFlash = Nothing
    End If
End If
End Sub

Private Sub mnuCopyReportWindow_Click()
frmReport.Caption = "Telegram/Server Output"
frmReport.AddReportLine rtbText.Text
frmReport.Show
End Sub

Private Sub rtbText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuBaseMenu
End Sub
