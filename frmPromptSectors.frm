VERSION 5.00
Begin VB.Form frmPromptSectors 
   Caption         =   "Select Sector(s)"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Optional Paramters"
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   3255
      Begin VB.TextBox txtParm 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select "
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cbField 
         Height          =   315
         ItemData        =   "frmPromptSectors.frx":0000
         Left            =   2040
         List            =   "frmPromptSectors.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optCond 
         Caption         =   "Select All F4 - Conditional Sectors"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtTerr 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton optTerr 
         Caption         =   "Territory"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtDest 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtSector 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optRange 
         Caption         =   "Range"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox txtRealm 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton OptRealm 
         Caption         =   "Realm #"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optSect 
         Caption         =   "Sector"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmPromptSectors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'092703 rjk: General reformatting
'101903 rjk: Modified to support Multiple Sectors Selection using the F4-Conditional Settings
'            Change to return a range if sectors are somewhat valid otherwise nothing for the range option
'            Modified to only update the calling form what you have something from the user.
'
'102503 rjk: Added support for the multiple territory fields (cbFields)
'111603 rjk: Changed the name TxtSector to txtSector, code cleanup.

Public HoldSource As Form
Public strSectors As String

Private Receiver As Control
Private Done As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Sub cmdOK_Click()
Dim strParms As String
Dim strSector As String

On Error Resume Next
strSector = vbNullString

'Get parms - pull off ? in front
strParms = txtParm.Text
If Len(Trim$(strParms)) > 0 Then
    If Left$(Trim$(strParms), 1) = "?" Then
        strParms = Mid$(Trim$(strParms), 2)
    End If
End If

'Single Sector
If optSect.Value Then
    strSector = txtSector.Text
    If Len(Trim$(txtParm.Text)) > 0 Then
        strSector = strSector + " ?" + strParms
    End If

'Realm
ElseIf OptRealm.Value Then
    strSector = "#" + txtRealm.Text
    If Len(Trim$(txtParm.Text)) > 0 Then
        strSector = strSector + " ?" + strParms
    End If

'Range
ElseIf optRange.Value Then
    Dim p1 As Integer
    Dim x1 As String
    Dim X2 As String
    Dim y1 As String
    Dim Y2 As String
    
    '101903 rjk: Change to return a range if sectors are somewhat valid otherwise nothing
    strSector = ""
    p1 = InStr(txtOrigin, ",")
    If p1 > 0 And Len(txtOrigin.Text) > p1 Then
        x1 = Trim$(Left$(txtOrigin, p1 - 1))
        y1 = Trim$(Mid$(txtOrigin, p1 + 1))
        p1 = InStr(txtDest, ",")
        If p1 > 0 And Len(txtDest.Text) > p1 Then
            X2 = Trim$(Left$(txtDest, p1 - 1))
            Y2 = Trim$(Mid$(txtDest, p1 + 1))
            strSector = x1 + ":" + X2 + "," + y1 + ":" + Y2
        End If
    End If
    If Len(Trim$(txtParm.Text)) > 0 Then
        strSector = strSector + " ?" + strParms
    End If
'Territory
ElseIf optTerr.Value Then
    '102503 rjk: Added support for the multiple territory fields
    If cbField.ListIndex = 0 Then
        strSector = "* ?terr=" + txtTerr.Text
    Else
        strSector = "* ?terr" + CStr(cbField.ItemData(cbField.ListIndex)) + "=" + txtTerr.Text
    End If
    'strSector = "* ?terr=" + txtTerr.Text
    If Len(Trim$(txtParm.Text)) > 0 Then
        strSector = strSector + "&" + strParms
    End If
'Conditional 101903 rjk: Added
ElseIf optCond.Value Then
    strSector = GetConditionalSectors()
End If

'Put the string in the Source

If Len(strSector) > 0 Then '101903 rjk: Update the field only you having information
    HoldSource.Enabled = True
    Receiver.Text = strSector
    'Receiver.SetFocus 101903 rjk: Not need to drawn focus to this field as we should be done with it
End If

Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next

HoldSource.Enabled = False

'101903 rjk: Added for Multiple Sector Selection using F4 conditional
If (Receiver.Name = "txtMultOrigin" Or Receiver.Name = "txtMultDest") And ColorScheme = COLORSCHEME_CONDITIONAL Then
    optCond.Enabled = True
End If

'If the range box is checked - make sure the correct boxes have the focus
If optRange.Value And Done = 0 Then
    If Len(txtOrigin.Text) = 0 Or Len(txtDest.Text) > 0 Then
        txtOrigin.SetFocus
    Else
        txtDest.SetFocus
    End If
    Done = 1
End If
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long    8/12/03 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

'Set to recieve focus for clicks on DrawMap
Set HoldSource = frmDrawMap.PromptForm
Set frmDrawMap.PromptForm = Me

Set Receiver = HoldSource.ActiveControl


'Fill Realm Combo
Dim X As Integer
For X = 1 To 20
    txtRealm.AddItem CStr(X)
Next X
txtRealm.ListIndex = 0
optRange.Value = True
cbField.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Done = 0
Set frmDrawMap.PromptForm = HoldSource
HoldSource.Enabled = True
Set HoldSource = Nothing
End Sub

Private Sub optRange_Click()
On Error Resume Next
txtOrigin.SetFocus
End Sub

Private Sub OptRealm_Click()
On Error Resume Next
txtRealm.SetFocus
End Sub

Private Sub optSect_Click()
On Error Resume Next
txtSector.SetFocus
End Sub

Private Sub optTerr_Click()
On Error Resume Next
txtTerr.SetFocus
End Sub

Private Sub txtDest_Change()
optRange.Value = True
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

On Error Resume Next
optRange.Value = True

If ParseSectors(sx, sy, txtOrigin.Text) Then
    txtDest.SetFocus
End If
End Sub

Private Sub txtRealm_Change()
OptRealm.Value = True
End Sub

Public Sub SetControls()
Dim Buffer As String
Dim parms As String
Dim x1 As String
Dim X2 As String
Dim y1 As String
Dim Y2 As String
Dim p1 As Integer
Dim p2 As Integer
Dim p3 As Integer

Buffer = Trim$(strSectors)
If Len(Buffer) = 0 Then
    Exit Sub
End If

'Remove Parameters
Dim X As Integer
X = InStr(Buffer, "?")
If X > 0 Then
    parms = Mid$(Buffer, X)
    txtParm.Text = Mid$(Buffer, X + 1)
    Buffer = Left$(Buffer, X - 1)
    
    'Check for territory
    If Mid$(parms, 2, 4) = "terr" Then
        p1 = InStr(6, parms, "&")
        If p1 > 0 Then
            txtTerr.Text = Mid$(parms, 7, p1 - 7)
            txtParm.Text = Mid$(parms, p1 + 1)
        Else
            txtTerr.Text = Mid$(parms, 7)
            txtParm.Text = vbNullString
        End If
    Exit Sub
    End If
End If

'check for Realm
If Left$(Buffer, 1) = "#" Then
    txtRealm.Text = Mid$(Buffer, 2)
    Exit Sub
End If


'Check for Range Entry
p1 = InStr(Buffer, ":")
If p1 > 0 Then
    x1 = Left$(Buffer, p1 - 1)
    p2 = InStr(p1 + 1, Buffer, ",")
    X2 = Mid$(Buffer, p1 + 1, p2 - p1 - 1)
    p3 = InStr(p2 + 1, Buffer, ":")
    y1 = Mid$(Buffer, p2 + 1, p3 - p2 - 1)
    Y2 = Mid$(Buffer, p3 + 1)
    txtOrigin.Text = x1 + "," + y1
    txtDest.Text = X2 + "," + Y2
    Exit Sub
End If

'Must Just be Single
txtSector.Text = Buffer
End Sub

Private Sub txtRealm_Click()
OptRealm.Value = True
End Sub

Private Sub txtSector_Change()
If Me.Visible = True Then
    optSect.Value = True
End If
End Sub

Private Sub txtTerr_Change()
optTerr = True
End Sub
