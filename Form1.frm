VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Splitter"
   ClientHeight    =   3465
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MousePointer    =   7  'Size N S
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   4125
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   2400
      MousePointer    =   1  'Arrow
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   2400
      MousePointer    =   1  'Arrow
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuStuff 
      Caption         =   "Stuff"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SPLITTER_HEIGHT = 40

' The percentage occupied by the PictureBox.
Private PercentageY As Single
Private PercentageX1 As Single
Private PercentageX2 As Single

' True when we are dragging the splitter.
Private Dragging As Boolean
Private DragY As Boolean
Private DragX1 As Boolean
' Arrange the controls on the form.
Private Sub ArrangeControls()
Dim hgt11 As Single
Dim hgt12 As Single
Dim hgt21 As Single
Dim hgt22 As Single
Dim wdt1 As Single
Dim wdt2 As Single

    ' Don't bother if we're iconized.
    If WindowState = vbMinimized Then Exit Sub

    wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
    wdt2 = (ScaleWidth - SPLITTER_HEIGHT) - wdt1
    hgt11 = (ScaleHeight - SPLITTER_HEIGHT) * PercentageX1
    hgt12 = (ScaleHeight - SPLITTER_HEIGHT) - hgt11
    hgt21 = (ScaleHeight - SPLITTER_HEIGHT) * PercentageX2
    hgt22 = (ScaleHeight - SPLITTER_HEIGHT) - hgt21
    
    Picture1.Move 0, 0, wdt1, hgt11
    Picture2.Move 0, hgt11 + SPLITTER_HEIGHT, wdt1, hgt12
    Picture3.Move wdt1 + SPLITTER_HEIGHT, 0, wdt2, hgt21
    Picture4.Move wdt1 + SPLITTER_HEIGHT, hgt21 + SPLITTER_HEIGHT, wdt2, hgt22
End Sub

Private Sub Form_Load()
    ' Start with the split in the middle.
    PercentageY = 0.5
    PercentageX1 = 0.5
    PercentageX2 = 0.5
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wdt1 As Single

    wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
    Dragging = True
    DragY = (X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT)
    DragX1 = (X < wdt1)
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim wdt1 As Single

If Not Dragging Then
    wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
    If X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT Then
        Me.MousePointer = vbSizeWE
    Else
        Me.MousePointer = vbSizeNS
    End If
Else
    If DragY Then
        PercentageY = X / ScaleWidth
        If PercentageY < 0 Then PercentageY = 0
        If PercentageY > 1 Then PercentageY = 1
    ElseIf DragX1 Then
        PercentageX1 = Y / ScaleHeight
        If PercentageX1 < 0 Then PercentageX1 = 0
        If PercentageX1 > 1 Then PercentageX1 = 1
    Else
        PercentageX2 = Y / ScaleHeight
        If PercentageX2 < 0 Then PercentageX2 = 0
        If PercentageX2 > 1 Then PercentageX2 = 1
    End If
    
    ArrangeControls
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub


Private Sub Form_Resize()
    ArrangeControls
End Sub

