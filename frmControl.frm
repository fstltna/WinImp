VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FileNumber As Integer
Dim j As Integer
Dim buf As String
Dim strDump() As String

FileNumber = FreeFile   ' Get unused file number.
Open App.Path + "\dump.dmp" For Input As #FileNumber

'Input the header
Line Input #FileNumber, buf
Line Input #FileNumber, buf
Line Input #FileNumber, buf
Line Input #FileNumber, buf

ReDim strDump(0 To 0)
j = 0

While Not EOF(FileNumber)
    Line Input #FileNumber, buf
    ReDim Preserve strDump(0 To j)
    strDump(j) = buf
    j = j + 1

Wend
Close #FileNumber

If j > 1 Then
    ReDim Preserve strDump(0 To j - 2)
    ProcessSectorDump strDump()
End If


End Sub

Private Sub Form_Load()
GameName = "Twisted Minds I"
OpenEmpireDB

End Sub
