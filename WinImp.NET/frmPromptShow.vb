Option Strict Off
Option Explicit On
Friend Class frmPromptShow
	Inherits System.Windows.Forms.Form
	
	' 08/11/03  efj     Commented everything out, nothing in this file is used
	
	'Private Sub cmdCancel_Click()
	'Unload Me
	'End Sub
	'
	'Private Sub cmdHelp_Click()
	'frmDrawMap.DisplayPromptHelp Label2.Caption
	'End Sub
	'
	'Public Sub cmdOk_Click()
	'Dim strCmd As String
	'Dim x As Integer
	'On Error Resume Next
	'
	''Start with first list box results
	'If Option1.Value Then
	'    strCmd = "land"
	'ElseIf Option2.Value Then
	'    strCmd = "plane"
	'ElseIf Option3.Value Then
	'    strCmd = "ship"
	'ElseIf Option7.Value Then
	'    strCmd = "bridge"
	'ElseIf Option9.Value Then
	'    strCmd = "tower"
	'ElseIf Option10.Value Then
	'    strCmd = "nuke"
	'Else
	'    strCmd = "sector"
	'End If
	'
	''Add result from second list box
	'If Option7.Value Or Option9.Value Or Option4.Value Then
	'    strCmd = strCmd + " build"
	'ElseIf Option5.Value Then
	'    strCmd = strCmd + " capacities"
	'Else
	'    strCmd = strCmd + " statistics"
	'End If
	'
	'strCmd = "show " + strCmd
	'
	''Add tech level if input
	'If Len(txtTechLevel.Text) > 0 Then
	'    x = val(txtTechLevel.Text)
	'    If x < 1 Then
	'        MsgBox "Tech String not valid", vbExclamation + vbOKOnly, "Entry Error"
	'        Exit Sub
	'    End If
	'    strCmd = strCmd + " " + CStr(val(txtTechLevel.Text))
	'End If
	'
	'frmEmpCmd.SubmitEmpireCommand strCmd, True
	'Unload Me
	'End Sub
	'
	'Private Sub Form_Load()
	''Make Stay always on top
	'Dim success As Long
	'success = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
	'
	'End Sub
	'
	'Private Sub Form_Unload(Cancel As Integer)
	'Set frmDrawMap.PromptForm = Nothing
	'frmDrawMap.PromptUp = False
	'End Sub
	'
	'Private Sub Option6_Click()
	'Dim n As Integer
	'If val(txtTechLevel.Text) < 1 And TechLevel > 0 Then
	'    n = CInt(TechLevel)
	'    If CSng(n) > TechLevel Then n = n - 1
	'    txtTechLevel.Text = CStr(n)
	'End If
	'End Sub
End Class