Option Strict Off
Option Explicit On
Friend Class frmAddUnitDes
	Inherits System.Windows.Forms.Form
	Public UnitClass As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		UnitClass = vbNullString
		Me.Hide()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		If Len(txtID.Text) = 0 Then
			MsgBox("Must enter a Unit id", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly)
			Exit Sub
		End If
		
		If Len(txtDesc.Text) = 0 Then
			MsgBox("Must enter a Description", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly)
			Exit Sub
		End If
		
		If Option1(0).Checked Then
			UnitClass = "s"
		ElseIf Option1(1).Checked Then 
			UnitClass = "l"
		ElseIf Option1(2).Checked Then 
			UnitClass = "p"
		End If
		
		Me.Hide()
	End Sub
End Class