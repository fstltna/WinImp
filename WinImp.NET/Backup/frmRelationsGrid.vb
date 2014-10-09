Option Strict Off
Option Explicit On
Friend Class frmRelationsGrid
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'093003 rjk: General reformatting.
	'112903 rjk: Added a Get news button and displaying of the telegram count in the grid.
	'121203 rjk: Fixed relations grid titles in Deity mode
	'180206 rjk: For 4.3.0 get country list, do not depend on relation.
	'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
	'            Switch nation relations to fill in xdump coun instead of relations.
	'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.
	
	Public Sub go()
		Dim X, Y As Short
		Dim strCmd As String
		
		If Me.Visible = False Then Exit Sub
		
		Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
		
		'okay, even after being reworked, the nationsList class is still a little
		'kludgy for what we want to do.  live with it for now drk@unxsoft.com
		
		MSFlexGrid1.Rows = Nations.ActiveCount + 1
		MSFlexGrid1.Cols = Nations.ActiveCount + 1
		
		Y = 1
		'121203 rjk: Fixed relations grid titles in Deity mode
		For X = 0 To Nations.Count
			If Nations.NationName(X) <> vbNullString Then
				MSFlexGrid1.set_TextMatrix(0, Y, Nations.NationName(X))
				MSFlexGrid1.set_TextMatrix(Y, 0, Nations.NationName(X))
				MSFlexGrid1.set_ColWidth(Y, 800)
				Y = Y + 1
			End If
		Next 
		
		If Y > 1 Then
			'found some active nations
			
			'100703 rjk: Moved the nation list building logic to Relations report
			'            must do our relations first as this resets the nation file.
			'            Also must skip otherwise the rsNations file will be reset.
			'            Reset the grid and request a refill.
			Nations.Clear()
			
			frmEmpCmd.SubmitEmpireCommand("db1") 'start database update
			If Not VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
				GetCountryList()
			End If
			If VersionCheck(4, 3, 4) = WinAceRoutines.enumVersion.VER_LESS Then
				frmEmpCmd.SubmitEmpireCommand("relation")
				
				rsNations.MoveFirst()
				While (Not rsNations.EOF)
					If rsNations.Fields("Number").Value <> CountryNumber Then
						strCmd = "relation " & rsNations.Fields("Number").Value
						frmEmpCmd.SubmitEmpireCommand(strCmd)
					End If
					rsNations.MoveNext()
				End While
			Else
				GetCountryList()
			End If
			frmEmpCmd.SubmitEmpireCommand("db2") 'finished with database update
			
		End If
	End Sub
	
	'well, this is a kludge, but there were no good callback hooks
	'so that I can know when the database update has completed, so I
	'just had to add mine own.
	
	'have frmEmpCmd call this every time it updates the nationsList object
	Public Sub callback()
		Dim Y, X, r As Short
		
		If Me.Visible = False Then Exit Sub
		
		MSFlexGrid1.row = 1
		MSFlexGrid1.col = 1
		MSFlexGrid1.Redraw = False
		
		For X = 0 To Nations.Count
			For Y = 0 To Nations.Count
				If (Nations.NationName(X) <> vbNullString And Nations.NationName(Y) <> vbNullString) Then
					r = Nations.Relation(Y, X)
					MSFlexGrid1.Text = Nations.RelationString(r)
					If (Nations.telegramCount(Y, X) > 0) And (Y <> X) Then
						MSFlexGrid1.Text = MSFlexGrid1.Text & " (" & CStr(Nations.telegramCount(Y, X)) & ")"
					End If
					Select Case r
						Case iREL_ALLIED : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 240, 0))
						Case iREL_FRIENDLY : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(160, 255, 160))
						Case iREL_NEUTRAL : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
						Case iREL_HOSTILE : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 150, 140))
						Case iREL_AT_WAR : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(240, 0, 0))
						Case iREL_MOBILIZING : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 100, 100))
						Case iREL_SITZKRIEG : MSFlexGrid1.CellBackColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 50, 50))
					End Select
					
					If (MSFlexGrid1.row = MSFlexGrid1.Rows - 1) Then
						MSFlexGrid1.row = 1
						If MSFlexGrid1.col < MSFlexGrid1.Cols - 1 Then MSFlexGrid1.col = MSFlexGrid1.col + 1
					Else
						MSFlexGrid1.row = MSFlexGrid1.row + 1
					End If
					
				End If
			Next 
		Next 
		
		MSFlexGrid1.Redraw = True
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Form property frmRelationsGrid.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		Me.Cursor = vbNormal
	End Sub
	
	Private Sub cmdGetNews_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGetNews.Click
		'Setup Designation dialog and execute
		frmDrawMap.PromptForm = frmPromptNews
		frmDrawMap.PromptUp = True
		
		'Put form in proper place
		frmPromptNews.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptNews.Width)) \ 2)
		frmPromptNews.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptNews.Height))
		
		'prompt for parameters
		frmPromptNews.Show()
	End Sub
	
	
	'UPGRADE_WARNING: Event frmRelationsGrid.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmRelationsGrid_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		MSFlexGrid1.Width = VB6.TwipsToPixelsX(Max(VB6.PixelsToTwipsX(Me.Width) - 300, 0))
		MSFlexGrid1.Height = VB6.TwipsToPixelsY(Max(VB6.PixelsToTwipsY(Me.Height) - 300, 0))
	End Sub
	
	Private Sub Slider1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Slider1.Change
		Dim X As Short
		
		For X = 0 To MSFlexGrid1.Cols - 1
			MSFlexGrid1.set_ColWidth(X, Slider1.Value * 10 + 10)
		Next 
	End Sub
End Class