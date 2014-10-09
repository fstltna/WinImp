VERSION 5.00
Begin VB.Form frmPromptThreshAll 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFoodRequired 
      Caption         =   "Use Food Required to determine Threshold"
      Height          =   735
      Left            =   3120
      TabIndex        =   36
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox chkSupply 
      Caption         =   "Auto Push/Pull supply to meet new thresholds"
      Height          =   735
      Left            =   4560
      TabIndex        =   35
      Top             =   1680
      Width           =   1455
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
      Left            =   5640
      TabIndex        =   34
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   13
      Left            =   5280
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   12
      Left            =   5280
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   11
      Left            =   5280
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   33
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   32
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   31
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   30
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   29
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   28
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   27
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   26
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   25
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   1320
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   20
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Set"
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
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "in"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Thresholds"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptThreshAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: commented out dead Sub SetThresh
'081103 efj: removal of dead variables
'090703 rjk: SetThresh is used in the ThresholdAll function (right click or Ctrl 't')
'091103 rjk: Change unc. workers to use word wrap.
'091803 rjk: Moved label setting functionality to SetLabels. Call SetLabels when origin changes
'092703 rjk: General reformatting.
'093003 rjk: switched to x and y derived from txtOrigin instead of SelX and SelY
'093003 rjk: Added to set initial field selection
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'240104 rjk: Renamed the Supply check, switched to item object, added multiple sector selection
'            Removed Field Transaction table, Update registry saving, added food required option,
'            Reorganized form.
'280104 rjk: Added double-click option.
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim arythresh(0 To 13) As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp "threshold"
End Sub

Public Sub cmdOK_Click()
Dim n As Integer
Dim strCmd As String
On Error Resume Next
Dim iComm As EmpItem
Dim lThresh As Long

For n = 0 To 13
    Set iComm = Items.FindByFormLabel(lblThresh(n).Caption)
    lThresh = Val(txtThresh(n).Text)
    If lThresh <> arythresh(n) Then
        frmPromptThresh.SetMultipleSectorThreshold iComm, txtMultOrigin.Text, lThresh, (chkFoodRequired <> vbUnchecked), (chkSupply <> vbUnchecked)
    End If
Next n

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

Private Sub Form_Activate()
txtMultOrigin.SetFocus '093003 rjk: Added to set initial field selection
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim sucess As Long   removed efj 8/2003
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 
'SetLabels '092103 rjk: Moved code to SetLabels.

'Put form in proper place
Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
top = frmDrawMap.top + frmDrawMap.Height - Height

If GetSetting(APPNAME, "frmPromptThresh", "Food Required", False) Then
    chkFoodRequired.Value = vbChecked
Else
    chkFoodRequired.Value = vbUnchecked
End If
If GetSetting(APPNAME, "frmPromptThresh", "Supply", False) Then
    chkSupply.Value = vbChecked
Else
    chkSupply.Value = vbUnchecked
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False

If chkFoodRequired.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptThresh", "Food Required", True
Else
    SaveSetting APPNAME, "frmPromptThresh", "Food Required", False
End If
If chkSupply.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptThresh", "Supply", True
Else
    SaveSetting APPNAME, "frmPromptThresh", "Supply", False
End If
End Sub

Private Sub txtMultOrigin_Change()
SetLabels
End Sub

Private Sub txtMultOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtMultOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub

Private Sub txtThresh_GotFocus(Index As Integer)
If Len(txtThresh(Index).Text) > 0 Then
    txtThresh(Index).SelStart = 0
    txtThresh(Index).SelLength = Len(txtThresh(Index).Text)
End If
End Sub

Private Sub SetLabels()
Dim i As Integer
Dim X As Integer
Dim Y As Integer

lblThresh(0).Caption = Items.FindByLetter("m").FormLabel
lblThresh(1).Caption = Items.FindByLetter("g").FormLabel
lblThresh(2).Caption = Items.FindByLetter("s").FormLabel
lblThresh(3).Caption = Items.FindByLetter("f").FormLabel
lblThresh(4).Caption = Items.FindByLetter("c").FormLabel
lblThresh(5).Caption = Items.FindByLetter("u").FormLabel
lblThresh(6).Caption = Items.FindByLetter("i").FormLabel
lblThresh(7).Caption = Items.FindByLetter("l").FormLabel
lblThresh(8).Caption = Items.FindByLetter("h").FormLabel
lblThresh(9).Caption = Items.FindByLetter("o").FormLabel
lblThresh(10).Caption = Items.FindByLetter("d").FormLabel
lblThresh(11).Caption = Items.FindByLetter("p").FormLabel
lblThresh(12).Caption = Items.FindByLetter("r").FormLabel
lblThresh(13).Caption = Items.FindByLetter("b").FormLabel

If ParseSectors(X, Y, txtMultOrigin.Text) And InStr(txtMultOrigin.Text, "\") = 0 Then
    rsSectors.Seek "=", X, Y
    
    For i = 0 To 13
        If Not rsSectors.NoMatch Then
            arythresh(i) = rsSectors.Fields(Items.FindByFormLabel(lblThresh(i).Caption).Letter + "_dist")
            txtThresh(i) = CStr(arythresh(i))
        Else
            arythresh(i) = 0
            txtThresh(i) = ""
        End If
    Next i
Else '093003 rjk: Otherwise blank values
    For i = 0 To 13
        arythresh(i) = 0
        txtThresh(i) = ""
    Next i
End If
End Sub
