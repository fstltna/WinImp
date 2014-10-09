Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmPrintMap
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'261104 rjk: Created
	'020605 rjk: Fixed the bmp file name.
	
	Private Sub cmdBmp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBmp.Click
		'NbrHexWide = Int(p.ScaleWidth / (HSL866 * 2))
		'NbrHexTall = Int(p.ScaleHeight / (HSL5 * 3))
		
		'frmPreview.ScaleMode = vbPixels
		'frmPreview.ScaleWidth = MapSizeX * HSL866 * 2
		'frmPreview.ScaleHeight = MapSizeY * (HSL5 * 3)
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form method frmPreview.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.Width = VB6.TwipsToPixelsX(frmPreview.ScaleX(MapSizeX * HSL866 * 2, vbPixels, vbTwips))
		'UPGRADE_ISSUE: Constant vbTwips was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form method frmPreview.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.Height = VB6.TwipsToPixelsY(frmPreview.ScaleY(MapSizeY * HSL5 * 3, vbPixels, vbTwips))
		'UPGRADE_ISSUE: Form property frmPreview.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.AutoRedraw = True
		'UPGRADE_ISSUE: Form method frmPreview.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.Cls()
		
		DrawGrid(frmPreview)
		'Printer.EndDoc
		'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		frmPreview.Image.Save("print_map.bmp")
		frmMDIPreview.Close()
		Me.Close()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdPreview_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreview.Click
		Dim Printer As New Printer
		Dim pPrinter As Printer
		Dim sMagnification As Single
		Dim iSaveOriginX As Short
		Dim iSaveOriginY As Short
		Dim iCenterX As Short
		Dim iCenterY As Short
		
		sMagnification = 1#
		
		For	Each pPrinter In Printers
			If VB6.GetItemString(cbPrinter, cbPrinter.SelectedIndex) = pPrinter.DeviceName Then
				Printer = pPrinter
				Exit For
			End If
		Next pPrinter
		Printer.ScaleMode = ScaleModeConstants.vbPixels
		'frmPreview.Width = Printer.ScaleWidth
		'frmPreview.Height = Printer.ScaleHeight
		frmPreview.Width = VB6.TwipsToPixelsX(Printer.Width)
		frmPreview.Height = VB6.TwipsToPixelsY(Printer.Height)
		'UPGRADE_ISSUE: Form property frmPreview.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.AutoRedraw = True
		'UPGRADE_ISSUE: Form method frmPreview.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.Cls()
		
		iSaveOriginX = OriginX
		iSaveOriginY = OriginY
		On Error Resume Next
		sMagnification = CSng(txtScale.Text)
		SetHexSideLength(frmPreview, (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * sMagnification)
		ParseSectors(iCenterX, iCenterY, txtOrigin.Text)
		CenterMap(frmPreview, iCenterX, iCenterY)
		DrawGrid(frmPreview)
		OriginX = iSaveOriginX
		OriginY = iSaveOriginY
		'SavePicture frmPreview.Image, "test.bmp"
		SetHexSideLength((frmDrawMap.picMap), (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * frmDrawMap.Magnification)
		Me.Close()
	End Sub
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
		Dim Printer As New Printer
		Dim pPrinter As Printer
		Dim sMagnification As Single
		Dim iSaveOriginX As Short
		Dim iSaveOriginY As Short
		Dim iCenterX As Short
		Dim iCenterY As Short
		
		sMagnification = 1#
		
		For	Each pPrinter In Printers
			If VB6.GetItemString(cbPrinter, cbPrinter.SelectedIndex) = pPrinter.DeviceName Then
				Printer = pPrinter
				Exit For
			End If
		Next pPrinter
		'Printer.ScaleMode = vbMillimeters
		'DrawGrid Printer
		Printer.ScaleMode = ScaleModeConstants.vbPixels
		frmPreview.Width = VB6.TwipsToPixelsX(Printer.Width)
		frmPreview.Height = VB6.TwipsToPixelsY(Printer.Height)
		'UPGRADE_ISSUE: Form property frmPreview.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.AutoRedraw = True
		'UPGRADE_ISSUE: Form method frmPreview.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmPreview.Cls()
		
		iSaveOriginX = OriginX
		iSaveOriginY = OriginY
		On Error Resume Next
		sMagnification = CSng(txtScale.Text)
		SetHexSideLength(frmPreview, (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * sMagnification)
		ParseSectors(iCenterX, iCenterY, txtOrigin.Text)
		CenterMap(frmPreview, iCenterX, iCenterY)
		DrawGrid(frmPreview)
		OriginX = iSaveOriginX
		OriginY = iSaveOriginY
		SetHexSideLength((frmDrawMap.picMap), (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX) / (30 * 0.866 * 2) * frmDrawMap.Magnification)
		'UPGRADE_ISSUE: PrintForm Component might require to be declared. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6D42BDC5-2EFB-44F3-942C-EAB9A2DB05F0"'
		frmPreview.PrintForm1.Print(frmPreview, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
		'Printer.EndDoc
		frmMDIPreview.Close()
		Me.Close()
	End Sub
	
	Private Sub frmPrintMap_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Printer As New Printer
		Dim pPrinter As Printer
		
		cbPrinter.SelectedIndex = -1
		cbPrinter.Items.Clear()
		For	Each pPrinter In Printers
			cbPrinter.Items.Add(pPrinter.DeviceName)
			If Printer.DeviceName = pPrinter.DeviceName Then
				cbPrinter.SelectedIndex = cbPrinter.Items.Count - 1
			End If
		Next pPrinter
		If cbPrinter.SelectedIndex = -1 Then
			cbPrinter.SelectedIndex = 0
		End If
		txtScale.Text = CStr(1#)
		txtOrigin.Text = "0,0"
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, txtOrigin.Text) Then
			cmdPreview.Enabled = True
			cmdPrint.Enabled = True
		Else
			cmdPreview.Enabled = False
			cmdPrint.Enabled = False
		End If
	End Sub
End Class