VERSION 5.00
Begin VB.Form frmPrintMap 
   Caption         =   "Print Map"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBmp 
      Caption         =   "bmp"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox chkSea 
      Caption         =   "Show Known Sea Sectors"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtOrigin 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   855
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "Landscape"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Portrait"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.ComboBox cbPrinter 
      Height          =   315
      ItemData        =   "frmPrintMap.frx":0000
      Left            =   720
      List            =   "frmPrintMap.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblCenter 
      Alignment       =   1  'Right Justify
      Caption         =   "Center"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblScale 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPrinter 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPrintMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'261104 rjk: Created
'020605 rjk: Fixed the bmp file name.

Private Sub cmdBmp_Click()
'NbrHexWide = Int(p.ScaleWidth / (HSL866 * 2))
'NbrHexTall = Int(p.ScaleHeight / (HSL5 * 3))

'frmPreview.ScaleMode = vbPixels
'frmPreview.ScaleWidth = MapSizeX * HSL866 * 2
'frmPreview.ScaleHeight = MapSizeY * (HSL5 * 3)
frmPreview.Width = frmPreview.ScaleX(MapSizeX * HSL866 * 2, vbPixels, vbTwips)
frmPreview.Height = frmPreview.ScaleY(MapSizeY * HSL5 * 3, vbPixels, vbTwips)
frmPreview.AutoRedraw = True
frmPreview.Cls

DrawGrid frmPreview
'Printer.EndDoc
SavePicture frmPreview.Image, "print_map.bmp"
Unload frmMDIPreview
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim pPrinter As Printer
Dim sMagnification As Single
Dim iSaveOriginX As Integer
Dim iSaveOriginY As Integer
Dim iCenterX As Integer
Dim iCenterY As Integer

sMagnification = 1#

For Each pPrinter In Printers
    If cbPrinter.List(cbPrinter.ListIndex) = pPrinter.DeviceName Then
        Set Printer = pPrinter
        Exit For
    End If
Next
Printer.ScaleMode = vbPixels
'frmPreview.Width = Printer.ScaleWidth
'frmPreview.Height = Printer.ScaleHeight
frmPreview.Width = Printer.Width
frmPreview.Height = Printer.Height
frmPreview.AutoRedraw = True
frmPreview.Cls

iSaveOriginX = OriginX
iSaveOriginY = OriginY
On Error Resume Next
sMagnification = CSng(txtScale)
SetHexSideLength frmPreview, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * sMagnification
ParseSectors iCenterX, iCenterY, txtOrigin
CenterMap frmPreview, iCenterX, iCenterY
DrawGrid frmPreview
OriginX = iSaveOriginX
OriginY = iSaveOriginY
'SavePicture frmPreview.Image, "test.bmp"
SetHexSideLength frmDrawMap.picMap, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * frmDrawMap.Magnification
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim pPrinter As Printer
Dim sMagnification As Single
Dim iSaveOriginX As Integer
Dim iSaveOriginY As Integer
Dim iCenterX As Integer
Dim iCenterY As Integer

sMagnification = 1#

For Each pPrinter In Printers
    If cbPrinter.List(cbPrinter.ListIndex) = pPrinter.DeviceName Then
        Set Printer = pPrinter
        Exit For
    End If
Next
'Printer.ScaleMode = vbMillimeters
'DrawGrid Printer
Printer.ScaleMode = vbPixels
frmPreview.Width = Printer.Width
frmPreview.Height = Printer.Height
frmPreview.AutoRedraw = True
frmPreview.Cls

iSaveOriginX = OriginX
iSaveOriginY = OriginY
On Error Resume Next
sMagnification = CSng(txtScale)
SetHexSideLength frmPreview, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * sMagnification
ParseSectors iCenterX, iCenterY, txtOrigin
CenterMap frmPreview, iCenterX, iCenterY
DrawGrid frmPreview
OriginX = iSaveOriginX
OriginY = iSaveOriginY
SetHexSideLength frmDrawMap.picMap, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * frmDrawMap.Magnification
frmPreview.PrintForm
'Printer.EndDoc
Unload frmMDIPreview
Unload Me
End Sub

Private Sub Form_Load()
Dim pPrinter As Printer

cbPrinter.ListIndex = -1
cbPrinter.Clear
For Each pPrinter In Printers
    cbPrinter.AddItem pPrinter.DeviceName
    If Printer.DeviceName = pPrinter.DeviceName Then
        cbPrinter.ListIndex = cbPrinter.ListCount - 1
    End If
Next
If cbPrinter.ListIndex = -1 Then
    cbPrinter.ListIndex = 0
End If
txtScale = CStr(1#)
txtOrigin = "0,0"
End Sub

Private Sub txtOrigin_Change()
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, txtOrigin) Then
    cmdPreview.Enabled = True
    cmdPrint.Enabled = True
Else
    cmdPreview.Enabled = False
    cmdPrint.Enabled = False
End If
End Sub
