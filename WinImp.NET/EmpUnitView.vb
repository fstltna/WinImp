Option Strict Off
Option Explicit On
Friend Class EmpUnitView
	
	'Change Log:
	'112803 rjk: Created
	'120103 rjk: Added a check for unknown relations (offline) to the enemyinfo generation
	'120303 rjk: Moved friendlyUnit to frmOptions
	'010404 rjk: Switched to screeninfo to a type
	'170404 rjk: Fixed the displaying of the Units when some units are disabled.
	'140704 rjk: Added Expired Units.
	'290606 rjk: Added Our Nuke Units.
	
	Enum enumViewingUnits
		VU_OUR_SHIPS = 0
		VU_OUR_LAND_UNITS = 1
		VU_OUR_PLANES = 2
		VU_OUR_NUKES = 3
		VU_ENEMY_SHIPS = 4
		VU_ENEMY_LAND_UNITS = 5
		VU_ENEMY_PLANES = 6
		VU_ALLIED_SHIPS = 7
		VU_ALLIED_LAND_UNITS = 8
		VU_ALLIED_PLANES = 9
		VU_NEUTRAL_SHIPS = 10
		VU_NEUTRAL_LAND_UNITS = 11
		VU_NEUTRAL_PLANES = 12
		VU_EXPIRED_SHIPS = 13
		VU_EXPIRED_LAND_UNITS = 14
		VU_EXPIRED_PLANES = 15
		'if more units are added, need to change lastViewingUnit and array dimensions
		'0100404 rjk: need to set the size in the iUnitCount in tScreenInfo to match the num enumViewingUnits
	End Enum
	
	Private Const VU_NUM_UNITS As Double = enumViewingUnits.VU_EXPIRED_PLANES + 1
	Private bViewUnit(VU_NUM_UNITS - 1) As Boolean
	Private vuPriorityList(VU_NUM_UNITS - 1) As enumViewingUnits
	Private vuPriorityListCount As Short
	Private picArray(VU_NUM_UNITS - 1) As System.Drawing.Image
	Private m_iExpiry As Short
	
	Public ReadOnly Property UnitPicture(ByVal iUnitCounts() As Short, ByVal bFirstUnit As Boolean) As System.Drawing.Image
		Get
			Static passCount As Short
			Static priorityListIndex As Short
			Dim Index As Short
			Dim modValue As Short
			Dim Value As Short
			
			If bFirstUnit Then
				passCount = 1
				priorityListIndex = -1
			End If
			
			For Index = 0 To vuPriorityListCount - 1
				priorityListIndex = priorityListIndex + 1
				If priorityListIndex >= priorityListCount Then
					priorityListIndex = 0
					passCount = passCount + 1
				End If
				If iUnitCounts(vuPriorityList(priorityListIndex)) >= passCount Then
					UnitPicture = picArray(vuPriorityList(priorityListIndex))
					Exit Property
				End If
			Next Index
			'UPGRADE_NOTE: Object UnitPicture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			UnitPicture = Nothing
		End Get
	End Property
	
	
	Public Property bViewUnits(ByVal Index As enumViewingUnits) As Boolean
		Get
			bViewUnits = bViewUnit(Index)
		End Get
		Set(ByVal Value As Boolean)
			bViewUnit(Index) = Value
		End Set
	End Property
	
	Public ReadOnly Property lastViewingUnit() As enumViewingUnits
		Get
			lastViewingUnit = VU_NUM_UNITS - 1
		End Get
	End Property
	
	
	Public Property iExpiry() As Short
		Get
			iExpiry = m_iExpiry
		End Get
		Set(ByVal Value As Short)
			m_iExpiry = Value
		End Set
	End Property
	
	
	Public Property priorityList(ByVal Index As enumViewingUnits) As enumViewingUnits
		Get
			priorityList = vuPriorityList(Index)
		End Get
		Set(ByVal Value As enumViewingUnits)
			vuPriorityList(Index) = Value
		End Set
	End Property
	
	
	Public Property priorityListCount() As Short
		Get
			priorityListCount = vuPriorityListCount
		End Get
		Set(ByVal Value As Short)
			vuPriorityListCount = Value
		End Set
	End Property
	
	Public Sub Save()
		Dim Index As enumViewingUnits
		Dim strTemp As String
		
		For Index = 0 To vuPriorityListCount - 1
			strTemp = strTemp & CStr(vuPriorityList(Index)) & ","
		Next Index
		If Len(strTemp) > 0 Then
			strTemp = Left(strTemp, Len(strTemp) - 1)
		End If
		SaveSetting(APPNAME, "UnitView", "PriorityList", strTemp)
		SaveSetting(APPNAME, "UnitView", "ExpiryDate", CStr(m_iExpiry))
	End Sub
	
	Public Sub Load()
		Dim Index As enumViewingUnits
		Dim strPriorityItems() As String
		Dim strTemp As String
		
		LoadUnitPictures()
		
		vuPriorityListCount = 0
		
		strTemp = GetSetting(APPNAME, "UnitView", "PriorityList", vbNullString)
		If Len(strTemp) = 0 Then
			For Index = 0 To lastViewingUnit
				bViewUnit(Index) = True
				vuPriorityList(Index) = Index
				vuPriorityListCount = vuPriorityListCount + 1
			Next Index
		Else
			For Index = 0 To lastViewingUnit
				bViewUnit(Index) = False
			Next Index
			strPriorityItems = Split(strTemp, ",")
			For Index = 0 To UBound(strPriorityItems)
				bViewUnit(CShort(strPriorityItems(Index))) = True
				vuPriorityList(Index) = CShort(strPriorityItems(Index))
				vuPriorityListCount = vuPriorityListCount + 1
			Next Index
		End If
		m_iExpiry = CShort(GetSetting(APPNAME, "UnitView", "ExpiryDate", CStr(0)))
	End Sub
	
	Private Sub LoadUnitPictures()
		On Error GoTo Empire_Error
		picArray(enumViewingUnits.VU_OUR_LAND_UNITS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\land.gif")
		picArray(enumViewingUnits.VU_OUR_SHIPS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\ship.gif")
		picArray(enumViewingUnits.VU_OUR_PLANES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\plane.gif")
		picArray(enumViewingUnits.VU_OUR_NUKES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\nuke.gif")
		picArray(enumViewingUnits.VU_ENEMY_LAND_UNITS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\eland.gif")
		picArray(enumViewingUnits.VU_NEUTRAL_LAND_UNITS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\nland.gif")
		picArray(enumViewingUnits.VU_ALLIED_LAND_UNITS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\fland.gif")
		picArray(enumViewingUnits.VU_EXPIRED_LAND_UNITS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\xland.gif")
		picArray(enumViewingUnits.VU_ENEMY_SHIPS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\eship.gif")
		picArray(enumViewingUnits.VU_NEUTRAL_SHIPS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\nship.gif")
		picArray(enumViewingUnits.VU_ALLIED_SHIPS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\fship.gif")
		picArray(enumViewingUnits.VU_EXPIRED_SHIPS) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\xship.gif")
		picArray(enumViewingUnits.VU_ENEMY_PLANES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\eplane.gif")
		picArray(enumViewingUnits.VU_NEUTRAL_PLANES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\nplane.gif")
		picArray(enumViewingUnits.VU_ALLIED_PLANES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\fplane.gif")
		picArray(enumViewingUnits.VU_EXPIRED_PLANES) = System.Drawing.Image.FromFile(My.Application.Info.DirectoryPath & "\xplane.gif")
		Exit Sub
		
Empire_Error: 
		EmpireError("LoadUnitPictures", vbNullString, vbNullString)
	End Sub
	
	Public Sub ClearAllUnits()
		Dim Index As enumViewingUnits
		
		For Index = 0 To lastViewingUnit
			bViewUnit(Index) = False
		Next Index
		vuPriorityListCount = 0
	End Sub
	
	Public Function VisibleUnits(ByRef iUnitCounts() As Short) As Boolean
		Dim Index As Short
		Dim modValue As Short
		Dim Value As Short
		
		For Index = 0 To vuPriorityListCount - 1
			If iUnitCounts(vuPriorityList(Index)) > 0 Then
				VisibleUnits = True
				Exit Function
			End If
		Next Index
		VisibleUnits = False
	End Function
End Class