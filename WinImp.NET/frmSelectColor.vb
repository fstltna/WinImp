Option Strict Off
Option Explicit On
Friend Class frmSelectColor
	Inherits System.Windows.Forms.Form
	'UPGRADE_WARNING: Lower bound of array HldC was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim HldC(NUMBER_OF_PLAYER_COLORS) As Integer
	'UPGRADE_WARNING: Lower bound of array HldT was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Dim HldT(NUMBER_OF_PLAYER_COLORS) As Integer
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Dim n As Short
		
		'Restore Colors
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			PlayerColors(n) = HldC(n)
			PlayerText(n) = HldT(n)
		Next n
		
		Me.Close()
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim n As Short
		
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			PlayerColors(n) = System.Drawing.ColorTranslator.ToOle(Picture3(n - 1).BackColor)
			PlayerText(n) = System.Drawing.ColorTranslator.ToOle(Picture3(n - 1).ForeColor)
		Next n
		frmDrawMap.DrawHexes()
		Me.Close()
	End Sub
	
	Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
		SetDefaultPlayerColors()
		FillColors()
	End Sub
	
	Private Sub cmdTest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTest.Click
		
		Dim n As Short
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			PlayerColors(n) = System.Drawing.ColorTranslator.ToOle(Picture3(n - 1).BackColor)
			PlayerText(n) = System.Drawing.ColorTranslator.ToOle(Picture3(n - 1).ForeColor)
		Next n
		frmDrawMap.DrawHexes()
		
	End Sub
	
	Private Sub frmSelectColor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Save Colors
		Dim n As Short
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			HldC(n) = PlayerColors(n)
			HldT(n) = PlayerText(n)
		Next n
		
		FillColors()
		'Center form on screen
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) \ 2)
		
		FillColors()
	End Sub
	
	Private Sub Picture3_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Picture3.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = Picture3.GetIndex(eventSender)
		If Button = VB6.MouseButtonConstants.LeftButton Then
			cd1Color.Color = System.Drawing.ColorTranslator.FromOle(PlayerColors(Index + 1))
			cd1Color.ShowDialog()
			PlayerColors(Index + 1) = System.Drawing.ColorTranslator.ToOle(cd1Color.Color)
		Else
			cd1Color.Color = System.Drawing.ColorTranslator.FromOle(PlayerText(Index + 1))
			cd1Color.ShowDialog()
			PlayerText(Index + 1) = System.Drawing.ColorTranslator.ToOle(cd1Color.Color)
		End If
		FillColors()
	End Sub
	
	Private Sub FillColors()
		Dim n As Short
		Dim ch As String
		
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			Picture3(n - 1).BackColor = System.Drawing.ColorTranslator.FromOle(PlayerColors(n))
		Next n
		
		For n = 1 To NUMBER_OF_PLAYER_COLORS
			'UPGRADE_ISSUE: PictureBox property Picture3.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture3(n - 1).AutoRedraw = True
			Picture3(n - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(PlayerText(n))
			SetDrawingFont(Picture3(n - 1))
			Picture3(n - 1).Font = VB6.FontChangeSize(Picture3(n - 1).Font, 10)
			ch = CStr(n)
			'UPGRADE_ISSUE: PictureBox method Picture3.TextWidth was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property Picture3.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture3(n - 1).CurrentX = (VB6.PixelsToTwipsX(Picture3(n - 1).ClientRectangle.Width) - Picture3(n - 1).TextWidth(ch)) / 2
			'UPGRADE_ISSUE: PictureBox method Picture3.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property Picture3.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture3(n - 1).CurrentY = (VB6.PixelsToTwipsY(Picture3(n - 1).ClientRectangle.Height) - Picture3(n - 1).TextHeight(ch)) / 2
			'UPGRADE_ISSUE: PictureBox method Picture3.Print was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			Picture3(n - 1).Print(ch)
		Next n
	End Sub
End Class