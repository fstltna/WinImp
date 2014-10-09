VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmToolCheMarker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Che/Plague Tool"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlague 
      Caption         =   "Mark Plague"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3135
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmToolCheMarker.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Mark Che"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Telegram"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmToolCheMarker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'070604 rjk: Added Plague Marking

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strSource As String
Dim strTemp As String
Dim n As Long

'loop through the telegram and mark che events
strSource = Text1.Text
Do
    n = InStr(strSource, vbLf)
    If n = 0 Then n = Len(strSource) + 1
    strTemp = Left$(strSource, n - 1)
    strSource = Mid$(strSource, n + 1)
    EventChe strTemp
Loop Until Len(strSource) = 0

Unload Me
End Sub

Private Sub cmdPlague_Click()
Dim strSource As String
Dim strTemp As String
Dim n As Long

'loop through the telegram and mark che events
strSource = Text1.Text
Do
    n = InStr(strSource, vbLf)
    If n = 0 Then n = Len(strSource) + 1
    strTemp = Left$(strSource, n - 1)
    strSource = Mid$(strSource, n + 1)
    EventPlague strTemp
Loop Until Len(strSource) = 0

Unload Me
End Sub
Private Sub Form_Load()
LoadList
If List1.ListCount > 0 Then
    List1.ListIndex = 0
End If
End Sub

Private Sub LoadList()

Dim hldindex
hldindex = rsTeleHead.Index
rsTeleHead.Index = "Received"
If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
    rsTeleHead.MoveLast
    While Not rsTeleHead.BOF
        If rsTeleHead.Fields("Class") = TELEGRAM_CLASS_PRODUCTION Then
            List1.AddItem rsTeleHead.Fields("Subject")
            List1.ItemData(List1.NewIndex) = rsTeleHead.Fields("ID")
        End If
        rsTeleHead.MovePrevious
    Wend
    rsTeleHead.MoveLast
End If

rsTeleHead.Index = hldindex

End Sub

Private Sub List1_Click()

'Load Text box with Telegram
Dim teleID As Integer
teleID = List1.ItemData(List1.ListIndex)
rsTeleBody.Index = "PrimaryKey"
rsTeleBody.Seek "=", teleID, 1
Text1.Text = vbNullString
'Load telegram body into text box
If Not rsTeleBody.NoMatch Then
    Do While (Not rsTeleBody.EOF)
        If rsTeleBody.Fields("ID") <> teleID Then
            Exit Do
        End If
        Text1.Text = Text1.Text + rsTeleBody.Fields("Body") + vbLf
        rsTeleBody.MoveNext
    Loop
End If

End Sub

