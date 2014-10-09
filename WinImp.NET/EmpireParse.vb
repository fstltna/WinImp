Option Strict Off
Option Explicit On
Module EmpireParse
	
	'Change Log:
	'092803 rjk: Fixed the parsing of ndump strings so capabilities
	'            and statistics would show correctly and could be used
	'            for computepathcost calculations.
	'100603 rjk: Changed to be mil/eff/tech = "0" to "???" to be consisted with
	'            the other fields where the information is not known in ParseLook and ParseSpyReport.
	'101503 rjk: Sometimes the nation is present with sub info and sometimes not during recon with ASW.
	'            Reorganized the code to accomodate present/absent of the nation/number and spaces
	'            in the name.  Also change code accomodate land sectors during ASW recon.
	'102103 rjk: The names come from the show sector s / rsSector table now.
	'            This logic needs to be redone for matching designations in ParseAttack and ParseLook
	'102503 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
	'102603 rjk: Server bug no space between End Sector and 'has arrived' in sorder(ParseShipOrder).
	'            Also occurs for 'no route possible' and 'route too long'.  Added code for parsing
	'            'no route possible' and 'route too long' to sorder(ParseShipOrder).
	'110403 rjk: Added support for Planes to ParseSpy routing.
	'110603 rjk: sorder on a sector gives an error
	'113003 rjk: Fixed the ParseLook to be able to correctly parse the number unit when a ship name has a "#" in it.
	'121103 rjk: Added a check to ensure we not picking up an abort from the intercept plane
	'121203 rjk: Switched to Items and Item objects.  Added parsing for show item.
	'240104 rjk: Fixed parsing problem with a sub seeing subs.
	'200204 rjk: Added check for # in the ship name in ParseMission
	'            Added check for ship name in order output, find using the blanks at the beginning
	'060304 rjk: Changed CSng to Val for regional settings
	'130304 rjk: ParseLook uses the implied knowledge to zero either military or civilians.
	'270304 rjk: Added 'Rollover Avail' to Version table
	'130604 rjk: Fill in a default for Unit description.
	'220305 rjk: Added 'Version' to Version table
	'140505 rjk: Fixed ParseCoastWatch to deal with # in the ship name.
	'160705 rjk: Update map if the current sector is being updated with sea values.
	'111105 rjk: Modified the show sect c to allow I_NONE product definition.
	'181205 rjk: Added max. eff. gain limits for sect.
	'281205 rjk: Added ParseShowInfrastructure.
	'291205 rjk: Fixed a bug where BIG CITY / GOREW can not be reset.
	'210106 rjk: Fixed "has a sail path" for sorder.
	'210106 rjk: Fixed show ship stats to work for opt_FUEL.
	'030206 rjk: Automatically refresh the unit grid when a mission changes.
	'180206 rjk: Fix ParseVersion to deal 4.3.0 new version format.
	'            Parse the research column for show nukes.
	'            Change item to eiItem.
	'140506 rjk: Added South Pacific Atlantis additional infrastructures.
	'220506 rjk: Fixed parsing for No Capital, does not have period anymore.
	'110706 rjk: Accomodate the change to show sect s for 4.3.6 servers.
	'181006 rjk: Added timestamps for Sea information.
	'020907 rjk: Extract TimeZoneAdj from ParseVersion.  Removing requirement
	'            for the update as it will be removed in a future version of
	'            server.  Store in the database as once the version value is
	'            determined xdump version does not see the timezone corrected
	'            time.
	'090907 rjk: Fix a crash when do a 'show product t' command.
	'131207 rjk: Add End cargo type
	'010108 rjk: Update Order form if it is open when get new order information.
	'220508 rjk: Delete enemy unit when it is blown up.
	'140608 rjk: Fixed a Data Conversion crash in ParseAndDeleteUnits, incorrect index
	'250309 rjk: Fixed ParseShowSector, format changes.
	
	'Parse the results of a coastwatch command
	Public Function ParseCoastWatch(ByRef strLook As String) As Object
		'      Springa (# 29) bbc  battlecruiser (#2418) @ -45,-21
		
		Dim SecList() As String
		Dim n As Short
		Dim n2 As Short
		' Dim n3 As Integer    8/2003 efj  removed
		' Dim t1 As Integer    8/2003 efj  removed
		Dim t2 As Short
		Dim sx As Short
		Dim sy As Short
		Dim strSect As String
		Dim owner As String
		
		On Error GoTo ParseCoastWatch_Exit
		
		'make sure sector list is at least size one
		ReDim SecList(ES_NUM)
		For n = 1 To ES_NUM
			SecList(n) = "???"
		Next n
		
		If Len(strLook) = 0 Then
			ParseCoastWatch = VB6.CopyArray(SecList)
			Exit Function
		End If
		
		'get the sector
		n = InStr(strLook, "@")
		If n > 0 Then
			strSect = Trim(Mid(strLook, n + 1))
			ParseSectors(sx, sy, strSect)
		Else
			ParseCoastWatch = VB6.CopyArray(SecList)
			Exit Function
		End If
		
		'get the owner
		n = InStr(strLook, "(#")
		If n > 0 Then
			n2 = InStr(n + 2, strLook, ")")
			t2 = n2 + 1
			owner = CStr(CShort(Trim(Mid(strLook, n + 2, n2 - n - 2))))
		End If
		
		'Unit designation
		SecList(EU_TYPE) = Trim(Mid(strLook, t2, 6))
		If InStr(SecList(EU_TYPE), " ") > 0 Then
			SecList(EU_TYPE) = Left(SecList(EU_TYPE), InStr(SecList(EU_TYPE), " ") - 1)
		End If
		
		'Unit number
		'Note - this is in the format  "(#z) - Ships
		n = InStr(t2, strLook, "(#")
		If n > 0 Then
			n = n + 2
		Else
			n = InStr(t2, strLook, "#") + 1
		End If
		n2 = InStr(n, strLook, ")") 'ships are followed by )
		If n2 = 0 Then
			n2 = InStr(n, strLook, "@") 'subs are followed by @
		End If
		If n2 > 0 Then
			SecList(EU_ID) = Trim(Mid(strLook, n, n2 - n))
		End If
		
		SecList(0) = GetUnitClass(CStr(SecList(EU_TYPE)), strLook)
		If Len(CStr(SecList(0))) = 0 Then
			Exit Function
		End If
		
		SecList(EU_OWNER) = owner
		SecList(EU_X) = CStr(sx)
		SecList(EU_Y) = CStr(sy)
		
		ParseCoastWatch = VB6.CopyArray(SecList)
		
		Exit Function
		
		'Error handling routine
ParseCoastWatch_Exit: 
		EmpireError("ParseCoastWatch", vbNullString, strLook)
		On Error Resume Next
		ParseCoastWatch = VB6.CopyArray(SecList)
	End Function
	'Parse the results of a sonar command
	Public Function ParseSonar(ByRef strSonar As String) As Object
		
		'sb   submarine (#12) @ 6,4
		'WinACE2 sb   submarine (#13) @ 6,4
		
		Dim n As Short
		Dim n2 As Short
		Dim NatNum As Short
		n = 0
		
		On Error GoTo Empire_Error
		
		'Sonar gives a reading similar to coastwatch
		'just need to add the country string
		NatNum = -1
		'101503 rjk: Sometimes the nation is present with sub info and sometimes not during recon with ASW.
		'            Reorganized to key on the separation between unit and description and try to
		'            accomodate spacing the nation name.
		'            DEPENDS on the ship unit names being shorter than 4 characters, I believe
		'While NatNum < 0
		'    n = InStr(n + 1, strSonar, " ")
		'    If n = 0 Then
		'        NatNum = 0
		'    Else
		'        NatNum = Nations.NationNumber(Left$(strSonar, n - 1))
		'    End If
		'Wend
		
		n2 = InStr(strSonar, "  ") 'separator between unit name and unit descriptor.
		If n2 = 0 Then '240104 rjk: Added an additional check on pattern matching
			NatNum = -1
		Else
			n = InStr(Left(strSonar, n2 - 2), " ")
			If n = 0 Then 'no nation name present
				NatNum = -1
			Else
				Do While InStr(n + 1, Left(strSonar, n2 - 2), " ") > 0 'Step to last space
					n = InStr(n + 1, Left(strSonar, n2 - 2), " ")
				Loop 
				NatNum = Nations.NationNumber(Left(strSonar, n - 1))
			End If
		End If
		strSonar = Left(strSonar, n) & "(#" & CStr(NatNum) & ") " & Mid(strSonar, n + 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseCoastWatch(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseSonar = ParseCoastWatch(strSonar)
		
		Exit Function
Empire_Error: 
		EmpireError("ParseSonar", vbNullString, strSonar)
		
	End Function
	
	'Parse for a number - used in telegrams
	Public Function ParseForNumber(ByRef strSample As String) As Short
		
		On Error GoTo Empire_Error
		If InStr(strSample, "a new") > 0 Then
			ParseForNumber = 1
		ElseIf InStr(strSample, " one ") > 0 Then 
			ParseForNumber = 1
		ElseIf InStr(strSample, " two ") > 0 Then 
			ParseForNumber = 2
		ElseIf InStr(strSample, " three ") > 0 Then 
			ParseForNumber = 3
		ElseIf InStr(strSample, " four ") > 0 Then 
			ParseForNumber = 4
		ElseIf InStr(strSample, " five ") > 0 Then 
			ParseForNumber = 5
		ElseIf InStr(strSample, " six ") > 0 Then 
			ParseForNumber = 6
		ElseIf InStr(strSample, " seven ") > 0 Then 
			ParseForNumber = 7
		ElseIf InStr(strSample, " eight ") > 0 Then 
			ParseForNumber = 8
		ElseIf InStr(strSample, " nine ") > 0 Then 
			ParseForNumber = 9
		ElseIf InStr(strSample, " ten ") > 0 Then 
			ParseForNumber = 10
		ElseIf InStr(strSample, " eleven ") > 0 Then 
			ParseForNumber = 11
		ElseIf InStr(strSample, " twelve ") > 0 Then 
			ParseForNumber = 12
		ElseIf InStr(strSample, " thirteen ") > 0 Then 
			ParseForNumber = 13
		ElseIf InStr(strSample, " fourteen ") > 0 Then 
			ParseForNumber = 14
		ElseIf InStr(strSample, " fifteen ") > 0 Then 
			ParseForNumber = 15
		ElseIf InStr(strSample, " sixteen ") > 0 Then 
			ParseForNumber = 16
		ElseIf InStr(strSample, " seventeen ") > 0 Then 
			ParseForNumber = 17
		ElseIf InStr(strSample, " eighteen ") > 0 Then 
			ParseForNumber = 18
		ElseIf InStr(strSample, " nineteen ") > 0 Then 
			ParseForNumber = 19
		Else
			ParseForNumber = 0
		End If
		
		If InStr(strSample, " twenty") > 0 Then
			ParseForNumber = ParseForNumber + 20
		ElseIf InStr(strSample, " thirty") > 0 Then 
			ParseForNumber = ParseForNumber + 30
		ElseIf InStr(strSample, " forty") > 0 Then 
			ParseForNumber = ParseForNumber + 40
		ElseIf InStr(strSample, " fifty") > 0 Then 
			ParseForNumber = ParseForNumber + 50
		ElseIf InStr(strSample, " sixty") > 0 Then 
			ParseForNumber = ParseForNumber + 60
		ElseIf InStr(strSample, " seventy") > 0 Then 
			ParseForNumber = ParseForNumber + 70
		ElseIf InStr(strSample, " eighty") > 0 Then 
			ParseForNumber = ParseForNumber + 80
		ElseIf InStr(strSample, " ninty") > 0 Then 
			ParseForNumber = ParseForNumber + 90
		ElseIf InStr(strSample, "several") > 0 Then 
			ParseForNumber = -1
		End If
		
		Exit Function
Empire_Error: 
		EmpireError("ParseForNumber", vbNullString, strSample)
		
	End Function
	Public Sub ParseVersion(ByRef strLine As String, ByRef rs As DAO.Recordset)
		Dim n As Short
		Dim n2 As Short
		Dim n3 As Short
		Dim strTokens() As String
		Dim strTemp As String
		Dim var As String
		Static OptionCheck As Boolean
		
		'Note - these routines use "val" rather than "CSng" for conversion, since
		'CSng is internationally aware, and the empire server ALWAYS uses
		'periods for decimal points.  Some countries setting expect a comma for
		'CSng to work correctly.
		
		On Error GoTo ParseVersion_Exit
		
		n = InStr(strLine, "Empire")
		n2 = InStr(strLine, "Wolfpack Empire")
		If n = 1 Or n2 = 1 Then
			var = "Major"
			n = InStr(strLine, ".")
			If n > 0 Then
				n2 = InStrRev(Left(strLine, n), " ")
				If n2 > 0 Then
					rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2 + 1, n - n2 - 1)))
					var = "Minor"
					n2 = InStr(n + 1, strLine, ".")
					If n2 > 0 Then
						rs.Fields(var).Value = CShort(Trim(Mid(strLine, n + 1, n2 - n - 1)))
						var = "Revision"
						n3 = InStr(n2 + 1, strLine, " ")
						If n3 > 0 Then
							rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2 + 1, n3 - n2 - 1)))
						Else
							rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2 + 1, Len(strLine) - n2)))
						End If
					End If
				End If
			End If
		End If
		
		'Check for world size
		n = InStr(strLine, "World size")
		If n > 0 Then
			n = InStr(strLine, "is")
			n2 = InStr(strLine, "by")
			n3 = InStr(strLine, ".")
			If n > 0 And n2 > 0 And n3 > 0 Then
				var = "World X"
				rs.Fields(var).Value = CShort(Trim(Mid(strLine, n + 3, n2 - n - 3)))
				var = "World Y"
				rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2 + 3, n3 - n2 - 3)))
			End If
			'300304 rjk: Ensure it is zero, if it is changed to 0 on the server, it will be missing from the version report
			rs.Fields("Rollover Avail").Value = 0
		End If
		
		'use The 'show' command to find out the time of the next update.
		'The current time is Sun Sep 02 14:47:08.
		'An update consists of 60 empire time units.
		
		'Check for time zone adjustment - difference between the local machine
		'game server
		n = InStr(strLine, "The current time is")
		If n > 0 Then
			var = "Time Zone Adj"
			'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
			rs.Fields(var).Value = DateDiff(Microsoft.VisualBasic.DateInterval.Minute, Now, EmpireTimeString("    " & Trim(Mid(strLine, 24, 15)) & " " & Str(Year(Now))))
		End If
		
		'Check for ETU's per update
		n = InStr(strLine, "An update consists")
		If n > 0 Then
			var = "ETU per update"
			n2 = 23
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		
		'Check for ETU length
		n = InStr(strLine, "An Empire time unit")
		If n > 0 Then
			var = "Empire Time Unit"
			n2 = 24
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = CShort(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Check for civs to produce a btu
		n = InStr(strLine, "produce a BTU")
		If n > 0 Then
			var = "Civs for BTUs"
			n2 = 10
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Non aggi food
		n = InStr(strLine, "A non-aggi")
		If n > 0 Then
			var = "Food grow"
			n2 = 43
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'harvest amount
		n = InStr(strLine, "will harvest")
		If n > 0 Then
			var = "Food harvest"
			n2 = 29
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Civilian birth rate
		n = InStr(strLine, "civilians will give birth")
		If n > 0 Then
			var = "Civ Birthrate"
			n2 = 35
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Unc. worker birth rate
		n = InStr(strLine, "workers will give birth")
		If n > 0 Then
			var = "Uw Birthrate"
			n2 = 47
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'People eat
		n = InStr(strLine, "people eat")
		If n > 0 Then
			var = "Adult food"
			n2 = 35
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
				rs.Fields("Food Needed").Value = "Y"
			End If
		End If
		
		
		'Babies eat
		n = InStr(strLine, "babies eat")
		If n > 0 Then
			var = "Baby food"
			n2 = 17
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'look for "No food is needed!!" - if found this is a blitz
		n = InStr(strLine, "No food is needed")
		If n > 0 Then
			var = "Food Needed"
			rs.Fields(var).Value = "N"
		End If
		
		'Interest rate
		n = InStr(strLine, "Banks pay")
		If n > 0 Then
			var = "Bank Interest"
			n2 = 12
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		
		'tax rate
		n = InStr(strLine, "civilians generate")
		If n > 0 Then
			n = InStr(strLine, "$")
			n2 = InStr(n + 1, strLine, ",")
			If n > 0 And n2 > 0 Then
				var = "Civ Taxes"
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n)))
				
				'Now get unskilled workers tax rate
				n = InStr(n2, strLine, "$")
				n2 = InStr(n + 1, strLine, " ")
				If n > 0 And n2 > 0 Then
					var = "UW Taxes"
					rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n)))
				End If
			End If
		End If
		
		'military salaries
		n = InStr(strLine, "active military cost")
		If n > 0 Then
			n = InStr(strLine, "$")
			n2 = InStr(n + 1, strLine, ",")
			If n > 0 And n2 > 0 Then
				var = "Mil Salary"
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n)))
				
				'Now get reserve cost tax rate
				n = InStr(n2, strLine, "$")
				n2 = InStr(n + 1, strLine, ".")
				n2 = InStr(n2 + 1, strLine, ".") 'We want the second period
				If n > 0 And n2 > 0 Then
					var = "Reserve Salary"
					rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n - 1)))
				End If
			End If
		End If
		
		'270304 rjk: Retro - ROLLOVER_SCALE - Up to 200 avail can roll over an update.
		n = InStr(strLine, " avail can roll over an update")
		If n > 0 Then
			var = "Rollover Avail"
			rs.Fields(var).Value = Val(Trim(Mid(strLine, 7, n - 7)))
		End If
		
		'Happiness
		n = InStr(strLine, "stroller")
		If n > 0 Then
			var = "Civs per stroller"
			n2 = 46
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Education
		n = InStr(strLine, "graduates")
		If n > 0 Then
			var = "Civs per graduate"
			n2 = 50
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		
		'Happiness average
		n = InStr(strLine, "Happiness is averaged")
		If n > 0 Then
			var = "Happy average"
			n2 = 28
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'Education average
		n = InStr(strLine, "Education is averaged")
		If n > 0 Then
			var = "Edu average"
			n2 = 28
			n = InStr(n2, strLine, " ")
			If n > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n2, n - n2)))
			End If
		End If
		
		'tech growth
		n = InStr(strLine, "logarithmic growth")
		If n > 0 Then
			var = "Tech base"
			n = InStr(strLine, "base") + 4
			n2 = InStr(n, strLine, ")")
			If n > 0 And n2 > 0 Then
				rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n - 1)))
				
				'Now get easy tech rate
				var = "Easy tech"
				n = InStr(strLine, "after")
				If n > 0 Then ' make sure there is an easy tech
					n = n + 5
					n2 = InStr(n + 1, strLine, ".")
					n2 = InStr(n2 + 1, strLine, ".") 'We want the second period
					If n > 0 And n2 > 0 Then
						rs.Fields(var).Value = Val(Trim(Mid(strLine, n + 1, n2 - n - 1)))
					End If
				End If
			End If
		End If
		
		'Tech decline
		'Nation levels (tech etc.) decline 1% every 96 time units.
		n = InStr(strLine, "(tech etc.) decline")
		If n > 0 Then
			var = "Tech Decay Rate"
			rs.Fields(var).Value = Val(Trim(StringBetween(strLine, "decline", "%")))
			var = "Tech Decay Time"
			rs.Fields(var).Value = Val(Trim(StringBetween(strLine, "every", "time units")))
		End If
		
		'Fire range
		n = InStr(strLine, "Fire ranges")
		If n > 0 Then
			var = "Fire range"
			n2 = 27
			rs.Fields(var).Value = Val(Trim(Mid(strLine, n2)))
		End If
		
		'Mobility gains
		n = InStr(strLine, "Maximum mobility")
		If n > 0 Then
			strTemp = strLine
			
			'Strings contain tabs to seperate instead of spaces
			n = InStr(strTemp, vbTab)
			While n > 0
				strTemp = Left(strTemp, n - 1) & " " & Mid(strTemp, n + 1)
				n = InStr(strTemp, vbTab)
			End While
			
			ParseDumpString(strTemp, strTokens)
			var = "Max mob - Sector"
			rs.Fields(var).Value = CShort(Trim(strTokens(2)))
			var = "Max mob - Ship"
			rs.Fields(var).Value = CShort(Trim(strTokens(3)))
			var = "Max mob - Plane"
			rs.Fields(var).Value = CShort(Trim(strTokens(4)))
			var = "Max mob - Land"
			rs.Fields(var).Value = CShort(Trim(strTokens(5)))
		End If
		
		'Mobility gains
		n = InStr(strLine, "Max eff gain")
		If n > 0 Then
			strTemp = strLine
			
			'Strings contain tabs to seperate instead of spaces
			n = InStr(strTemp, vbTab)
			While n > 0
				strTemp = Left(strTemp, n - 1) & " " & Mid(strTemp, n + 1)
				n = InStr(strTemp, vbTab)
			End While
			
			ParseDumpString(strTemp, strTokens)
			If Trim(strTokens(5)) <> "--" Then
				var = "Eff gain - Sect"
				rs.Fields(var).Value = CShort(Trim(strTokens(5)))
			End If
			var = "Eff gain - Ship"
			rs.Fields(var).Value = CShort(Trim(strTokens(6)))
			var = "Eff gain - Plane"
			rs.Fields(var).Value = CShort(Trim(strTokens(7)))
			var = "Eff gain - Land"
			rs.Fields(var).Value = CShort(Trim(strTokens(8)))
		End If
		
		'Mobility increase
		n = InStr(strLine, "Max mob gain")
		If n > 0 Then
			strTemp = strLine
			
			'Strings contain tabs to seperate instead of spaces
			n = InStr(strTemp, vbTab)
			While n > 0
				strTemp = Left(strTemp, n - 1) & " " & Mid(strTemp, n + 1)
				n = InStr(strTemp, vbTab)
			End While
			
			ParseDumpString(strTemp, strTokens)
			var = "Mob gain - Sector"
			rs.Fields(var).Value = CShort(Trim(strTokens(5)))
			var = "Mob gain - Ship"
			rs.Fields(var).Value = CShort(Trim(strTokens(6)))
			var = "Mob gain - Plane"
			rs.Fields(var).Value = CShort(Trim(strTokens(7)))
			var = "Mob gain - Land"
			rs.Fields(var).Value = CShort(Trim(strTokens(8)))
		End If
		
		'Option Check
		n = InStr(strLine, "Options enabled in this game")
		If n > 0 Then
			var = "BIG_CITY"
			rs.Fields(var).Value = False
			var = "GO_RENEW"
			rs.Fields(var).Value = False
			OptionCheck = True
		End If
		
		'Options enabled in this game:
		'        ALL_BLEED, BIG_CITY, BLITZ, DEMANDUPDATE, FALLOUT, GODNEWS, GO_RENEW,
		'        INTERDICT_ATT, LANDSPIES, LOANS, MARKET, NEUTRON, NEW_STARVE,
		'        NEW_WORK, NEWPOWER, NO_PLAGUE, NOFOOD, NOMOBCOST, NUKEFAILDETONATE,
		'        ORBIT, PINPOINTMISSILE, PLANENAMES, SAIL, SHIPNAMES, SHOWPLANE,
		'        SUPER_BARS , TREATIES, UPDATESCHED'
		
		'Options disabled in this game:
		'        BRIDGETOWERS, DEFENSE_INFRA, DRNUKE, EASY_BRIDGES, FUEL, GRAB_THINGS,
		'        HIDDEN, LOSE_CONTACT, MOB_ACCESS, NO_FORT_FIRE, NO_HCMS, NO_LCMS,
		'        NO_OIL, NONUKES, RES_POP, SHIP_DECAY, SLOW_WAR, SNEAK_ATTACK,
		'        TRADESHIPS,
		
		'Option Check
		n = InStr(strLine, "Options disabled in this game")
		If n > 0 Then
			OptionCheck = False
		End If
		
		If OptionCheck Then
			n = InStr(strLine, "BIG_CITY")
			If n > 0 Then
				var = "BIG_CITY"
				rs.Fields(var).Value = True
			End If
			
			n = InStr(strLine, "GO_RENEW")
			If n > 0 Then
				var = "GO_RENEW"
				rs.Fields(var).Value = True
			End If
		End If
		
		Exit Sub
		
		'Error handling routine
ParseVersion_Exit: 
		On Error Resume Next
		EmpireError("ParseVersion", var, strLine)
		
	End Sub
	Public Sub ParseNation(ByRef strLine As String)
		Dim n As Short
		Dim n2 As Short
		' Dim n3 As Integer    8/2003 efj  removed
		' Dim strTokens() As String    8/2003 efj  removed
		' Dim strTemp As String    8/2003 efj  removed
		Dim var As String
		
		'(#1) 1 Nation Report    Sat Dec 12 15:51:28 1998
		'Nation status is ACTIVE     Bureaucratic Time Units: 640
		'100% eff capital at 0,0 has 999 civilians & 55 military
		'No capital. (was at -18,4)
		' The treasury has $26400.00     Military reserves: 0
		'Education..........  0.00       Happiness.......  0.00
		'Technology.........112.15       Research........  9.44
		'Technology factor : 51.95%     Plague factor :   0.00%
		'
		'Max population : 999
		'Max safe population for civs/uws: 768/868
		'Happiness needed Is 1.803645
		
		'Note - these routines use "val" rather than "CSng" for conversion, since
		'CSng is internationally aware, and the empire server ALWAYS uses
		'periods for decimal points.  Some countries setting expect a comma for
		'CSng to work correctly.
		
		On Error GoTo ParseNation_Exit
		
		'Check for country number
		n = InStr(strLine, "Nation Report")
		If n > 0 Then
			var = "CountryNumber"
			'100703 rjk: removed not used
			n = InStr(strLine, "at") '100703 rjk: removed not used
			n2 = InStr(strLine, "has") '100703 rjk: removed not used
			CountryNumber = Int(CDbl(StringBetween(strLine, "(#", ")")))
			'100703 rjk: Added country name, to pickup on "change name" command
			frmEmpCmd.CountryName = Mid(strLine, InStr(strLine, ")") + 2)
			frmEmpCmd.CountryName = Left(frmEmpCmd.CountryName, InStr(frmEmpCmd.CountryName, "Nation Report") - 2)
			Nations.AddNation(CountryNumber, (frmEmpCmd.CountryName))
		End If
		
		'Check for captial number
		n = InStr(strLine, "eff capital")
		If n = 0 Then n = InStr(strLine, "eff mountain capital")
		If n > 0 Then
			var = "CapSect"
			n = InStr(strLine, "at")
			n2 = InStr(strLine, "has")
			If n > 0 And n2 > 0 Then
				CapSect = Trim(Mid(strLine, n + 2, n2 - n - 2))
			End If
		End If
		
		n = InStr(strLine, "No capital")
		If n > 0 Then
			var = "CapSect"
			CapSect = Trim(StringBetween(strLine, "at", ")"))
		End If
		
		'Check for Military reserves
		n = InStr(strLine, "Military reserves:")
		If n > 0 Then
			n = n + 18
			var = "MilReserves"
			MilReserves = Val(Trim(Mid(strLine, n)))
		End If
		
		'Check for education
		n = InStr(strLine, "Education")
		If n > 0 Then
			var = "Education"
			Education = Val(Trim(Mid(strLine, n + 19, 6)))
		End If
		
		'Check for tech level
		n = InStr(strLine, "Technology..")
		If n > 0 Then
			var = "TechLevel"
			TechLevel = Val(Trim(Mid(strLine, n + 19, 6)))
		End If
		
		'Check for Happiness
		n = InStr(strLine, "Happiness..")
		If n > 0 Then
			var = "Happiness"
			Happiness = Val(Trim(Mid(strLine, n + 16)))
		End If
		
		'Check for Research
		n = InStr(strLine, "Research")
		If n > 0 Then
			var = "Research"
			Research = Val(Trim(Mid(strLine, n + 16)))
		End If
		
		'Max population
		n = InStr(strLine, "Max population")
		If n > 0 Then
			var = "Maxpop"
			n = InStr(strLine, ":")
			Maxpop = Val(Trim(Mid(strLine, n + 1)))
		End If
		
		'Check for Research
		n = InStr(strLine, "Max safe")
		If n > 0 Then
			var = "MaxSafeCivs"
			n = InStr(strLine, ":")
			n2 = InStr(n, strLine, "/")
			MaxSafeCivs = Val(Trim(Mid(strLine, n + 1, n2 - n - 1)))
			var = "MaxSafeUws"
			MaxSafeUws = Val(Trim(Mid(strLine, n2 + 1)))
		End If
		
		
		'Check for Happiness needed
		n = InStr(strLine, "Happiness needed")
		If n > 0 Then
			var = "HapNeeded"
			HapNeeded = Val(Trim(Mid(strLine, n + 19)))
		End If
		
		Exit Sub
		
		'Error handling routine
ParseNation_Exit: 
		EmpireError("ParseNation", var, strLine)
		
	End Sub
	
	'Parse results of DUMP string to update database
	Public Sub ParseDumpString(ByRef strLine As String, ByRef strTokens() As String, Optional ByRef Deityflag As Boolean = False)
		
		'if this is a deity dump, include special handling for country code
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Deityflag) Then
			Deityflag = False
		End If
		
		Dim j As Short
		Dim n As Short
		Dim fstart As Short
		Dim strSpace As String
		Dim strQuote As String
		Dim country As String
		Dim NeedCountry As Boolean
		
		On Error GoTo Empire_Error
		
		ReDim strTokens(0)
		strSpace = " "
		strQuote = Chr(34)
		j = 0
		strLine = strLine & strSpace
		fstart = 1
		NeedCountry = True
		Do 
			'note - special handling is required for quotes, since ship
			'names can have internal spaces surrounded by quotes.
			If Mid(strLine, fstart, 1) = strQuote Then
				n = InStr(fstart + 1, strLine, strQuote) + 1
			Else
				n = InStr(fstart, strLine, strSpace)
			End If
			
			If n > fstart Then
				If Deityflag And NeedCountry Then 'save off country
					country = Mid(strLine, fstart, n - fstart)
					NeedCountry = False
				Else
					ReDim Preserve strTokens(j)
					strTokens(j) = Mid(strLine, fstart, n - fstart)
					j = j + 1
				End If
			End If
			
			fstart = n + 1
		Loop Until n = 0
		
		If Deityflag Then 'put country back on the end
			ReDim Preserve strTokens(j)
			strTokens(j) = country
		End If
		
		Exit Sub
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseDumpString", CStr(n), strLine)
		
	End Sub
	
	'Parse results of Look command
	Public Function ParseLook(ByRef strLook As String) As Object
		Dim SecList() As String
		Dim n As Short
		Dim n2 As Short
		Dim n3 As Short
		Dim t1 As Short
		Dim t2 As Short
		Dim strDes As String
		
		On Error GoTo ParseLook_Exit
		
		'make sure sector list is at least size one
		ReDim SecList(ES_NUM)
		For n = 1 To ES_NUM
			SecList(n) = "???"
		Next n
		
		'Ignore info on your own sectors
		If LCase(Left(strLook, 4)) = "your" Then
			ParseLook = VB6.CopyArray(SecList)
			Exit Function
		End If
		
		If Len(strLook) = 0 Then
			ParseLook = VB6.CopyArray(SecList)
			Exit Function
		End If
		
		Dim sx As Short
		Dim sy As Short
		Dim strSect As String
		Dim owner As String
		Dim Mil As String
		
		'get the owner
		n = InStr(strLook, "(#")
		If n > 0 Then
			n2 = InStr(n + 2, strLook, ")")
			t2 = n2 + 1
			owner = CStr(CShort(Mid(strLook, n + 2, n2 - n - 2)))
		End If
		
		'get the sector
		n = InStr(strLook, "@")
		If n > 0 Then
			strSect = Trim(Mid(strLook, n + 1))
			ParseSectors(sx, sy, strSect)
		End If
		
		'get the mil
		n = InStr(t2, strLook, " mil")
		If n > 0 Then
			n2 = n - 1
			While n2 > 0 And Mid(strLook, n2, 1) <> " "
				n2 = n2 - 1
			End While
			Mil = Trim(Mid(strLook, n2 + 1, n - n2 - 1))
		Else
			Mil = "???" '100603 rjk: Changed to be "???" to be consisted with the field where the information is not known.
		End If
		
		'get the efficency
		n = InStr(strLook, "%")
		If n = 0 Then
			SecList(EU_MIL) = Mil
			SecList(EU_OWNER) = owner
			SecList(EU_X) = CStr(sx)
			SecList(EU_Y) = CStr(sy)
			
			'Unit designation
			SecList(EU_TYPE) = Trim(Mid(strLook, t2, 6))
			'113003 rjk: Moved the unit number until after we know if it is a ship or not.
			'            Helps with parsing a ship with '#' in its name.
			'get the class (ship, plane, land)
			SecList(0) = GetUnitClass(CStr(SecList(EU_TYPE)), strLook)
			If Len(CStr(SecList(0))) = 0 Then
				Exit Function
				'Unit number
				'Note - this could be in the format "#z " (land and plane) or
				'                                   "(#z) - Ships
			ElseIf SecList(0) = "s" Then  'ship
				n = InStr(t2, strLook, "(#")
				n2 = InStr(n + 1, strLook, ")") 'ships are followed by )
				SecList(EU_ID) = Trim(Mid(strLook, n + 2, n2 - n - 2))
			Else 'land or plane
				n = InStr(t2, strLook, "#")
				n2 = InStr(n + 1, strLook, " ") 'land units or planes are followed by space
				SecList(EU_ID) = Trim(Mid(strLook, n + 1, n2 - n - 1))
			End If
		Else
			SecList(0) = "h"
			SecList(ES_MIL) = Mil
			SecList(ES_OWNER) = owner
			SecList(ES_X) = CStr(sx)
			SecList(ES_Y) = CStr(sy)
			
			'Sector efficency
			n2 = n - 1
			While n2 > 0 And Mid(strLook, n2, 1) <> " "
				n2 = n2 - 1
			End While
			t1 = n2
			SecList(ES_EFFICIENCY) = Trim(Mid(strLook, n2 + 1, n - n2 - 1))
			
			'get the civilian
			n = InStr(t1, strLook, " civ")
			If n > 0 Then
				n2 = n - 1
				While n2 > 0 And Mid(strLook, n2, 1) <> " "
					n2 = n2 - 1
				End While
				SecList(ES_CIV) = Trim(Mid(strLook, n2 + 1, n - n2 - 1))
				If SecList(ES_MIL) = "???" Then '130304 rjk: If civilians are found then there are no military
					SecList(ES_MIL) = "0"
				End If
			Else
				If SecList(ES_MIL) <> "???" Then '130304 rjk: If military are found then there are no civilians
					SecList(ES_CIV) = "0"
				Else
					SecList(ES_CIV) = "???"
				End If
			End If
			
			'hardest part - get the sector
			strDes = Trim(Mid(strLook, t2, t1 - t2))
			
			'Handle the custom cases
			'102103 rjk: The names come from the show sector s / rsSector table now.
			'            This logic needs to be redone.
			'If strDes = "mine" Then
			'    strDes = "iron mine"
			'ElseIf strDes = "technical center" Then
			'    strDes = "tech center"
			'ElseIf strDes = "bridge span" Then
			'    strDes = "e span"
			'ElseIf strDes = "bridge head" Then
			'    strDes = "e head"
			'End If
			
			'load the designation
			n = 1
			Do 
				'    If InStr(SectDesString(n), Left$(strDes, 6)) > 0 Then
				If InStr(SectDesString(n), strDes) > 0 And Len(SectDesString(n)) = Len(strDes) + 4 Then
					SecList(ES_DES) = Left(SectDesString(n), 1)
					Exit Do
				End If
				n = n + 1
			Loop Until Len(SectDesString(n)) = 0
		End If
		
		ParseLook = VB6.CopyArray(SecList)
		Exit Function
		
		'Error handling routine
ParseLook_Exit: 
		
		EmpireError("ParseLook", vbNullString, strLook)
		ParseLook = VB6.CopyArray(SecList)
	End Function
	
	'Parse results of Attack command
	Public Function ParseAttack(ByRef strLook As String) As Object
		Dim SecList() As String
		Dim n As Short
		' Dim n2 As Integer    8/2003 efj  removed
		Dim t1 As Short
		Dim t2 As Short
		Dim strDes As String
		
		On Error GoTo ParseAttack_Exit
		'-1,3 is a 0% Wolf highway with approximately 0 military.
		
		'make sure sector list is at least size one
		ReDim SecList(ES_NUM)
		For n = 1 To ES_NUM
			SecList(n) = "???"
		Next n
		
		If Len(strLook) = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseAttack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ParseAttack = vbNullString
			Exit Function
		End If
		
		Dim sx As Short
		Dim sy As Short
		Dim strSect As String
		Dim strTemp As String
		Dim strOwner As String
		Dim owner As Short
		Dim Mil As String
		
		'get the owner
		strOwner = vbNullString
		owner = -1
		t1 = InStr(strLook, "%") + 2
		t2 = 0
		strTemp = Mid(strLook, t1)
		Do While owner < 0
			t2 = InStr(t2 + 1, strTemp, " ")
			If t2 = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ParseAttack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ParseAttack = vbNullString
				Exit Function
				'drk@unxsoft.com.  the diety country doesn't seem to be in the nations list, so probing
				'a diety country doesn't enter anything into the database.  We should at least enter something...
				'fixme
			End If
			strOwner = Trim(Left(strTemp, t2 - 1))
			owner = Nations.NationNumber(strOwner)
		Loop 
		
		'get the sector
		n = InStr(strLook, " ")
		If n > 0 Then
			strSect = Trim(Left(strLook, n))
			ParseSectors(sx, sy, strSect)
		End If
		
		'get the mil
		Mil = CStr(Val(Trim(StringBetween(strLook, "approximately", "military"))))
		
		'get the efficency
		SecList(0) = "h"
		SecList(ES_MIL) = Mil
		SecList(ES_OWNER) = CStr(owner)
		SecList(ES_X) = CStr(sx)
		SecList(ES_Y) = CStr(sy)
		
		'Sector efficency
		SecList(ES_EFFICIENCY) = CStr(Val(Trim(StringBetween(strLook, "is a", "%"))))
		
		'hardest part - get the sector
		n = InStr(t1 + t2 + 1, strLook, " with ")
		
		strDes = Trim(Mid(strLook, t1 + t2, n - (t2 + t1)))
		
		'Handle the custom cases
		'102103 rjk: The names come from the show sector s / rsSector table now.
		'            This logic needs to be redone.
		'If strDes = "mine" Then
		'    strDes = "iron mine"
		'ElseIf strDes = "technical center" Then
		'    strDes = "tech center"
		'ElseIf strDes = "bridge span" Then
		'    strDes = "e span"
		'ElseIf strDes = "bridge head" Then
		'    strDes = "e head"
		'End If
		
		'load the designation
		n = 1
		Do 
			'    If InStr(SectDesString(n), Left$(strDes, 6)) > 0 Then
			If InStr(SectDesString(n), strDes) > 0 And Len(SectDesString(n)) = Len(strDes) + 4 Then
				SecList(ES_DES) = Left(SectDesString(n), 1)
				Exit Do
			End If
			n = n + 1
		Loop Until Len(SectDesString(n)) = 0
		
		ParseAttack = VB6.CopyArray(SecList)
		Exit Function
		
		'Error handling routine
ParseAttack_Exit: 
		EmpireError("ParseAttack", vbNullString, strLook)
		ParseAttack = VB6.CopyArray(SecList)
	End Function
	
	
	'Parse results of a Recon flight
	Public Function ParseReconReport(ByRef strLine As String) As Object
		
		Dim strTokens() As String
		Dim SecList() As String
		Dim numTokens As Short
		Dim n As Short
		
		Dim sx As Short
		Dim sy As Short
		' Dim strSect As String    8/2003 efj  removed
		' Dim owner As String    8/2003 efj  removed
		' Dim Mil As String    8/2003 efj  removed
		' Dim Eff As String    8/2003 efj  removed
		' Dim Tech As String    8/2003 efj  removed
		' Dim utype As String    8/2003 efj  removed
		' Dim sid As String    8/2003 efj  removed
		
		On Error GoTo ParseReconReport_Exit
		
		'make sure sector list is at least size one
		ReDim SecList(ES_NUM)
		For n = 1 To ES_NUM
			SecList(n) = "???"
		Next n
		
		If Mid(strLine, 5, 1) = "," Then
			
			ParseDumpString(strLine, strTokens)
			numTokens = UBound(strTokens)
			
			For n = LBound(strTokens) To numTokens
				Select Case n
					Case 0
						If ParseSectors(sx, sy, strTokens(n)) Then
							SecList(ES_X) = CStr(sx)
							SecList(ES_Y) = CStr(sy)
						End If
					Case 1
						SecList(ES_DES) = strTokens(n)
					Case 2
						SecList(ES_OWNER) = strTokens(n)
					Case 3
						SecList(ES_EFFICIENCY) = strTokens(n)
					Case 4
						SecList(ES_ROAD) = strTokens(n)
					Case 5
						SecList(ES_RAIL) = strTokens(n)
					Case 6
						SecList(ES_DEFENSE) = strTokens(n)
					Case 7
						SecList(ES_CIV) = strTokens(n)
					Case 8
						SecList(ES_MIL) = strTokens(n)
					Case 9
						SecList(ES_SHELL) = strTokens(n)
					Case 10
						SecList(ES_GUN) = strTokens(n)
					Case 11
						SecList(ES_IRON) = strTokens(n)
					Case 12
						SecList(ES_PETROL) = strTokens(n)
					Case 13
						SecList(ES_FOOD) = strTokens(n)
				End Select
				
			Next n
			
			'Skip if its on our country
			If CDbl(SecList(ES_OWNER)) = CountryNumber Then
				SecList(0) = vbNullString
			Else
				SecList(0) = "h"
			End If
			
			'Late percent sign means its a unit
			'Added a check to ensure we not picking up an aborted from the intercept plane
		ElseIf Right(Trim(strLine), 1) = "%" And Right(Trim(strLine), 2) <> " %" And InStr(strLine, "aborted") = 0 Then 
			ParseDumpString(Mid(strLine, 3), strTokens)
			numTokens = UBound(strTokens)
			
			For n = LBound(strTokens) To numTokens
				Select Case n
					Case 0
						SecList(EU_OWNER) = strTokens(n)
					Case 1
						SecList(EU_ID) = strTokens(n)
					Case 2
						SecList(EU_TYPE) = strTokens(n)
					Case numTokens - 1
						If ParseSectors(sx, sy, strTokens(n)) Then
							SecList(ES_X) = CStr(sx)
							SecList(ES_Y) = CStr(sy)
						End If
					Case numTokens
						SecList(EU_EFF) = Left(strTokens(n), Len(strTokens(n)) - 1)
				End Select
			Next n
			
			'Skip if its on our country
			If CDbl(SecList(EU_OWNER)) = CountryNumber Then
				SecList(0) = vbNullString
			Else
				'get the class (ship, plane, land)
				SecList(0) = GetUnitClass(CStr(SecList(EU_TYPE)), strLine)
				If Len(CStr(SecList(0))) = 0 Then
					Exit Function
				End If
			End If
			
			'check for country status change
		ElseIf InStr(strLine, " downgraded ") > 0 Then 
			'Country WinACEdv2 (#4) has downgraded their relations with you to "Hostile"!
			'Another cold war...
			'Diplomatic relations with WinACEdv2 downgraded to "Hostile".
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseReconReport. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ParseReconReport = VB6.CopyArray(SecList)
			Exit Function
			
		ElseIf InStr(strLine, "(#") > 0 And InStr(strLine, "@") > 0 Then 
			'fly (#8) highway 100% efficient with approx 1000 civ @ 6,10
			'but not Jaytopia    completely  car  aircraft carrier Enterprise(#219)
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseLook(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ParseReconReport = ParseLook(strLine)
			Exit Function
		End If
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseReconReport. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseReconReport = VB6.CopyArray(SecList)
		Exit Function
		
		'Error handling routine
ParseReconReport_Exit: 
		EmpireError("ParseReconReport", vbNullString, strLine)
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseReconReport. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseReconReport = VB6.CopyArray(SecList)
	End Function
	
	
	'Parse the results of a Spy report
	Public Function ParseSpyReport(ByRef strLine As String) As Object
		
		Dim strTokens() As String
		Dim SecList() As String
		Dim numTokens As Short
		Dim n As Short
		Dim n2 As Short
		
		Dim sx As Short
		Dim sy As Short
		Dim strSect As String
		Dim owner As String
		Dim Mil As String
		Dim Eff As String
		Dim Tech As String
		Dim utype As String
		Dim sid As String
		
		On Error GoTo ParseSpyReport_Exit
		
		'make sure sector list is at least size one
		ReDim SecList(ES_NUM)
		For n = 1 To ES_NUM
			SecList(n) = "???"
		Next n
		
		'New unit format
		'Neutral (disneyland) unit in 8,6:  cav  cavalry #81 (eff 95, mil 20, tech 199)
		'Enemy (1 2 3) plane in 18,4: sf   F-117A Nighthawk #27
		
		'If InStr(strLine, "Spies report") > 0 Then
		If InStr(strLine, "unit in") > 0 Then
			'get the owner
			n = InStr(strLine, "(")
			If n > 0 Then
				n2 = InStr(n + 1, strLine, ")")
				owner = CStr(Nations.NationNumber(Mid(strLine, n + 1, n2 - n - 1)))
			End If
			
			'get the sector
			n = InStr(strLine, "unit in")
			If n > 0 Then
				n = n + 7
				n2 = InStr(n, strLine, ":")
				strSect = Trim(Mid(strLine, n, n2 - n))
				ParseSectors(sx, sy, strSect)
			End If
			
			'Get the type
			utype = Trim(Mid(strLine, n2 + 3, 4))
			
			'get the unit id
			n = InStr(n2 + 1, strLine, "#")
			If n > 0 Then
				n = n + 1
				n2 = InStr(n, strLine, " ")
				If n2 > 0 Then
					sid = Trim(Mid(strLine, n, n2 - n))
				Else
					sid = "999"
				End If
			End If
			
			'get the mil
			n = InStr(strLine, ", mil ")
			If n > 0 Then
				n = n + 6
				n2 = InStr(n, strLine, ",")
				If n2 > n Then
					Mil = Trim(Mid(strLine, n, n2 - n))
				Else
					Mil = "0"
				End If
			Else
				Mil = "???" '100603 rjk: Changed to be "???" to be consisted with the other fields where the information is not known.
			End If
			
			'get the efficiency
			n = InStr(strLine, "(eff ")
			If n > 0 Then
				n = n + 5
				n2 = InStr(n, strLine, ",")
				Eff = Trim(Mid(strLine, n, n2 - n))
			Else
				Eff = "???" '100603 rjk: Changed to be "???" to be consisted with the other fields where the information is not known.
			End If
			
			'get the tech
			n = InStr(strLine, ", tech ")
			If n > 0 Then
				n = n + 7
				n2 = InStr(n, strLine, ")")
				Tech = Trim(Mid(strLine, n, n2 - n))
			Else
				Tech = "???" '100603 rjk: Changed to be "???" to be consisted with the other fields where the information is not known.
			End If
			
			SecList(EU_MIL) = Mil
			SecList(EU_TECH) = Tech
			SecList(EU_EFF) = Eff
			SecList(EU_OWNER) = owner
			SecList(EU_TYPE) = utype
			SecList(EU_X) = CStr(sx)
			SecList(EU_Y) = CStr(sy)
			SecList(EU_ID) = sid
			
			'get the class (ship, plane, land)
			SecList(0) = GetUnitClass(CStr(SecList(EU_TYPE)), strLine)
			If Len(CStr(SecList(0))) = 0 Then
				Exit Function
			End If
			'Enemy (1 2 3) plane in 18,4: sf   F-117A Nighthawk #27 110403 rjk: Added support for Planes.
		ElseIf InStr(strLine, "plane in") > 0 Then 
			'get the owner
			n = InStr(strLine, "(")
			If n > 0 Then
				n2 = InStr(n + 1, strLine, ")")
				owner = CStr(Nations.NationNumber(Mid(strLine, n + 1, n2 - n - 1)))
			End If
			
			'get the sector
			n = InStr(strLine, "plane in")
			If n > 0 Then
				n = n + 8
				n2 = InStr(n, strLine, ":")
				strSect = Trim(Mid(strLine, n, n2 - n))
				ParseSectors(sx, sy, strSect)
			End If
			
			'Get the type
			utype = Trim(Mid(strLine, n2 + 2, 4))
			
			'get the unit id
			n = InStr(n2 + 1, strLine, "#")
			If n > 0 Then
				n = n + 1
				sid = Mid(strLine, n)
			End If
			
			SecList(EU_OWNER) = owner
			SecList(EU_TYPE) = utype
			SecList(EU_X) = CStr(sx)
			SecList(EU_Y) = CStr(sy)
			SecList(EU_ID) = sid
			
			'get the class (ship, plane, land)
			SecList(0) = GetUnitClass(CStr(SecList(EU_TYPE)), strLine)
			If Len(CStr(SecList(0))) = 0 Then
				Exit Function
			End If
		Else
			'skip headers and comments
			If Not (Left(strLine, 1) = " " Or Left(strLine, 1) = "-") Or Left(strLine, 4) = "    " Or Left(strLine, 4) = "   s" Then
				ParseSpyReport = VB6.CopyArray(SecList)
				Exit Function
			End If
			SecList(0) = "h"
			
			ParseDumpString(strLine, strTokens)
			numTokens = UBound(strTokens)
			
			For n = LBound(strTokens) To numTokens
				Select Case n
					Case 0
						If ParseSectors(sx, sy, strTokens(n)) Then
							SecList(ES_X) = CStr(sx)
							SecList(ES_Y) = CStr(sy)
						End If
					Case 1
						SecList(ES_DES) = strTokens(n)
					Case 2
						SecList(ES_OWNER) = strTokens(n)
					Case 3
						SecList(ES_OLDOWNER) = strTokens(n)
					Case 4
						SecList(ES_EFFICIENCY) = strTokens(n)
					Case 5
						SecList(ES_ROAD) = strTokens(n)
					Case 6
						SecList(ES_RAIL) = strTokens(n)
					Case 7
						SecList(ES_DEFENSE) = strTokens(n)
					Case 8
						SecList(ES_CIV) = strTokens(n)
					Case 9
						SecList(ES_MIL) = strTokens(n)
					Case 10
						SecList(ES_SHELL) = strTokens(n)
					Case 11
						SecList(ES_GUN) = strTokens(n)
					Case 12
						SecList(ES_PETROL) = strTokens(n)
					Case 13
						SecList(ES_FOOD) = strTokens(n)
					Case 14
						SecList(ES_BARS) = strTokens(n)
				End Select
				
			Next n
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseSpyReport. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseSpyReport = VB6.CopyArray(SecList)
		Exit Function
		
		'Error handling routine
ParseSpyReport_Exit: 
		EmpireError("ParseSpyReport", vbNullString, strLine)
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseSpyReport. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseSpyReport = VB6.CopyArray(SecList)
	End Function
	'Parse results of show build string to update database
	Public Sub ParseShow(ByRef strLine As String, ByRef tokenstart As Short, ByRef UnitType As String, ByRef ReportType As String, ByRef LineNumber As Short)
		
		Static txtLineOne As String
		
		'save the first line, because we don't know the report type yet
		If LineNumber = 1 Then
			txtLineOne = strLine
			Exit Sub
		End If
		
		'Now, if we have a saved first line, output it
		If Len(txtLineOne) > 0 Then
			ParseShowText(txtLineOne, UnitType, ReportType, 1)
			txtLineOne = vbNullString
		End If
		
		If LineNumber <= 0 Then
			Exit Sub
		End If
		
		Select Case UnitType
			Case "l", "s", "p", "n"
				ParseShowBuild(strLine, tokenstart, UnitType, ReportType)
			Case "h"
				ParseShowSector(strLine, tokenstart, ReportType)
			Case "c"
				Select Case ReportType
					Case "s", "c", "h"
						ParseShowSector(strLine, tokenstart, ReportType)
					Case "b"
						ParseShowInfrastructure(strLine, LineNumber)
				End Select
			Case "b", "t"
				ParseShowText(strLine, UnitType, ReportType, LineNumber)
			Case "i"
				ParseShowItem(strLine, LineNumber)
		End Select
		
	End Sub
	
	'Parse results of show build string to update database
	Public Sub ParseShowBuild(ByRef strLine As String, ByRef tokenstart As Short, ByRef UnitClass As String, ByRef ReportType As String)
		Dim j As Short
		Dim k As Short
		Dim n As Short
		Dim fstart As Short
		Dim strSpace As String
		Dim strID As String
		Dim strDesc As String
		Dim strTokens() As String
		Dim strValues As String
		
		If tokenstart <= 0 Then
			Exit Sub
		End If
		
		On Error GoTo ParseShowBuild_Exit
		
		n = 1
		strID = Trim(Left(strLine, 5))
		n = 2
		strDesc = Trim(Mid(strLine, 6, tokenstart - 6))
		
		ReDim strTokens(0)
		strSpace = " "
		j = 0
		strValues = Trim(Mid(strLine, tokenstart)) & strSpace
		fstart = 1
		Do 
			n = InStr(fstart, strValues, strSpace)
			If n > fstart Then
				ReDim Preserve strTokens(j)
				strTokens(j) = Mid(strValues, fstart, n - fstart)
				j = j + 1
			End If
			
			fstart = n + 1
		Loop Until n = 0
		
		'update the build records
		rsBuild.Seek("=", UnitClass, strID)
		If rsBuild.NoMatch Then
			rsBuild.AddNew()
			'initialize fields
			For j = 0 To rsBuild.Fields.Count - 1
				If rsBuild.Fields(j).Type = DAO.DataTypeEnum.dbText Then
					rsBuild.Fields(j).Value = " "
				Else
					rsBuild.Fields(j).Value = 0
				End If
			Next j
			rsBuild.Fields("id").Value = strID
			rsBuild.Fields("type").Value = UnitClass
		Else
			rsBuild.Edit()
		End If
		
		'now fill in the various fields to update the database
		rsBuild.Fields("desc").Value = strDesc
		n = 2
		
		If ReportType = "b" Then
			rsBuild.Fields("lcm").Value = CShort(strTokens(0))
			n = 3
			rsBuild.Fields("hcm").Value = CShort(strTokens(1))
			If UnitClass = "s" Then
				n = 4
				rsBuild.Fields("avail").Value = CShort(strTokens(2))
				n = 5
				rsBuild.Fields("tech").Value = CShort(strTokens(3))
				n = 6
				rsBuild.Fields("cost").Value = CInt(Mid(strTokens(4), 2))
			ElseIf UnitClass = "n" Then 
				n = 4 '092503 rjk: switched to 4 from 5 to be consistent
				rsBuild.Fields("oil").Value = CShort(strTokens(2))
				n = 5 '092503 rjk: switched to 5 from 6 to be consistent
				rsBuild.Fields("rad").Value = CShort(strTokens(3))
				n = 6 '092503 rjk: switched to 6 from 7 to be consistent
				rsBuild.Fields("avail").Value = CShort(strTokens(4))
				n = 7 '092503 rjk: switched to 7 from 8 to be consistent
				rsBuild.Fields("tech").Value = CShort(strTokens(5))
				n = 8 '092503 rjk: switched to 8 from 9 to be consistent
				'check for a stand alone '$' on nukes
				If strTokens(7) = "$" Then
					rsBuild.Fields("stat5").Value = CShort(strTokens(6)) 'Research
					rsBuild.Fields("cost").Value = CInt(strTokens(8))
				Else
					rsBuild.Fields("stat5").Value = 0 'Research
					rsBuild.Fields("cost").Value = CInt(strTokens(7))
				End If
			Else
				n = 4
				If UnitClass = "p" Then
					rsBuild.Fields("mil").Value = CShort(strTokens(2))
				Else
					rsBuild.Fields("gun").Value = CShort(strTokens(2))
				End If
				
				n = 5
				rsBuild.Fields("avail").Value = CShort(strTokens(3))
				n = 6
				rsBuild.Fields("tech").Value = CShort(strTokens(4))
				n = 7
				rsBuild.Fields("cost").Value = CInt(Mid(strTokens(5), 2))
			End If
		ElseIf UnitClass = "n" Then 
			'092503 rjk: Switched to load values in stat1, stat2, etc.
			'n = 9
			'For j = LBound(strTokens) To UBound(strTokens)
			'    If n < 13 Then
			'        rsBuild.Fields(n).Value = CInt(strTokens(j))
			'    ElseIf n = 15 Then
			'        'abilities
			'        n = 29
			'        rsBuild.Fields(n).Value = Trim$(strTokens(j))
			'    ElseIf n > 29 And n < 49 Then
			'        rsBuild.Fields(n).Value = Trim$(strTokens(j))
			'    End If
			'    n = n + 1
			'Next j
			If UBound(strTokens) >= 5 Then
				n = 9
				rsBuild.Fields("stat1").Value = CShort(strTokens(0))
				n = 10
				rsBuild.Fields("stat2").Value = CShort(strTokens(1))
				n = 11
				rsBuild.Fields("stat3").Value = CShort(strTokens(2))
				n = 12
				rsBuild.Fields("stat4").Value = CShort(strTokens(3))
				'skip dollar sign
				If strTokens(4) <> "$" Then
					n = 13
					rsBuild.Fields("stat5").ValidateOnSet = CShort(strTokens(4))
					n = 15
					k = 6
				Else
					n = 14
					k = 5
				End If
				
				rsBuild.Fields("cost").Value = CInt(strTokens(k))
				For j = 1 To UBound(strTokens) - k
					'abilities
					n = k + j + 9
					rsBuild.Fields("cap" & CStr(j)).Value = Trim(strTokens(j + k))
				Next j
			Else 'invalid input
				GoTo ParseShowBuild_Exit
			End If
		ElseIf ReportType = "s" Then 
			n = 12
			For j = LBound(strTokens) To UBound(strTokens)
				If n < 32 Then
					If n < 14 And UnitClass = "l" Then
						'land attack & defense are single
						rsBuild.Fields(n).Value = Val(strTokens(j)) '060304 rjk: Changed to Val for regional settings
					ElseIf UnitClass = "p" And (InStr(strTokens(j), "%") > 0) Then 
						'plane steath has percentage sign
						rsBuild.Fields(n).Value = CShort(Left(strTokens(j), Len(strTokens(j)) - 1))
					Else
						'everything else is Int
						If InStr(strTokens(j), "/") = 0 Then
							rsBuild.Fields(n).Value = CShort(strTokens(j))
						End If
					End If
				End If
				n = n + 1
			Next j
		ElseIf ReportType = "c" Then 
			n = 32
			For j = LBound(strTokens) To UBound(strTokens)
				If n < 52 Then
					rsBuild.Fields(n).Value = Trim(strTokens(j))
					n = n + 1
				End If
			Next j
		End If
		
		rsBuild.Update()
		Exit Sub
		
		'Error handling routine
ParseShowBuild_Exit: 
		EmpireError("ParseShowBuild", CStr(n), strLine)
		
	End Sub
	
	'Parse results of show build string to update database
	Public Sub ParseShowSector(ByRef strLine As String, ByRef tokenstart As Short, ByRef ReportType As String)
		Dim j As Short
		Dim n As Short
		Dim fstart As Short
		Dim strSpace As String
		Dim strDes As String
		Dim strDesc As String
		Dim strTemp As String
		Dim strTokens() As String
		Dim strValues As String
		Static strHeader As String
		If ReportType = "h" Then
			strHeader = strLine & " "
			Exit Sub
		End If
		
		If tokenstart <= 0 Then
			Exit Sub
		End If
		
		On Error GoTo ParseShowSector_Exit
		
		n = 1
		strDes = Trim(Left(strLine, 1))
		n = 2
		strDesc = Trim(Mid(strLine, 3, tokenstart - 3))
		
		'update the build records
		rsSectorType.Seek("=", strDes)
		If rsSectorType.NoMatch Then
			rsSectorType.AddNew()
			'initialize fields
			For j = 0 To rsSectorType.Fields.Count - 1
				If rsSectorType.Fields(j).Type = DAO.DataTypeEnum.dbText Then
					rsSectorType.Fields(j).Value = " "
				Else
					rsSectorType.Fields(j).Value = 0
				End If
			Next j
			rsSectorType.Fields("des").Value = strDes
		Else
			rsSectorType.Edit()
		End If
		
		Dim eiItem As EmpItem
		Select Case ReportType
			Case "s"
				
				'parse into tokens
				ReDim strTokens(0)
				strSpace = " "
				j = 0
				strValues = Trim(Mid(strLine, tokenstart)) & strSpace
				fstart = 1
				Do 
					n = InStr(fstart, strValues, strSpace)
					If n > fstart Then
						ReDim Preserve strTokens(j)
						strTokens(j) = Mid(strValues, fstart, n - fstart)
						j = j + 1
					End If
					
					fstart = n + 1
				Loop Until n = 0
				
				'now fill in the various fields to update the database
				rsSectorType.Fields("desc").Value = strDesc
				If UBound(strTokens) = 9 Then
					For n = 0 To 9
						If n = 2 Or n = 3 Then
							rsSectorType.Fields(n + 1).Value = Val(strTokens(n))
						ElseIf n = 0 Then 
							If strTokens(n) <> "no" Then
								rsSectorType.Fields("mcost").Value = Val(strTokens(n))
							End If
						ElseIf n = 1 Then 
							If strTokens(n) <> "way" Then
								rsSectorType.Fields("mcost100").Value = Val(strTokens(n))
							End If
						ElseIf n = 4 Then 
							Select Case strTokens(n)
								Case "sea"
									rsSectorType.Fields("nav").Value = 1
								Case "land"
									rsSectorType.Fields("nav").Value = 0
								Case "harbor"
									rsSectorType.Fields("nav").Value = 2
								Case "canal"
									rsSectorType.Fields("nav").Value = 3
								Case "bridge"
									rsSectorType.Fields("nav").Value = 4
							End Select
							'navig
						ElseIf n = 5 Then 
							Select Case strTokens(n)
								Case "normal"
									rsSectorType.Fields("pack_type").Value = 1
								Case "urban"
									rsSectorType.Fields("pack_type").Value = 3
								Case "warehouse"
									rsSectorType.Fields("pack_type").Value = 2
								Case "bank"
									rsSectorType.Fields("pack_type").Value = 4
							End Select
						Else
							rsSectorType.Fields(Maxpop).Value = Val(strTokens(n))
						End If
					Next n
				ElseIf UBound(strTokens) = 6 Then 
					For n = 0 To 6
						If n = 2 Or n = 3 Then
							rsSectorType.Fields(n + 1).Value = Val(strTokens(n))
						ElseIf n = 0 Then 
							If strTokens(n) <> "no" Then
								rsSectorType.Fields(n + 2).Value = Val(strTokens(n))
							End If
						ElseIf n = 1 Then 
							If strTokens(n) <> "way" Then
								rsSectorType.Fields("mcost100").Value = Val(strTokens(n))
							End If
						Else
						End If
					Next 
					
				Else
					For n = 0 To 8
						If n = 1 Or n = 2 Then
							rsSectorType.Fields(n + 2).Value = Val(strTokens(n)) '060304 rjk: Changed to Val for regional settings.
						Else
							rsSectorType.Fields(n + 2).Value = CShort(strTokens(n))
						End If
					Next n
				End If
				
			Case "c"
				
				n = InStr(tokenstart, strHeader, "product")
				If n > 0 Then
					rsSectorType.Fields("product").Value = Trim(Mid(strLine, n, 7))
				End If
				
				n = InStr(tokenstart, strHeader, "use1")
				If n > 0 Then
					strTemp = Trim(Mid(strLine, n + 3, 1))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use1s").Value = strTemp
					Else
						rsSectorType.Fields("use1s").Value = vbNullString
					End If
					strTemp = Trim(Mid(strLine, n, 2))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use1n").Value = CShort(strTemp)
					Else
						rsSectorType.Fields("use1n").Value = 0
					End If
				End If
				
				n = InStr(tokenstart, strHeader, "use2")
				If n > 0 Then
					strTemp = Trim(Mid(strLine, n + 3, 1))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use2s").Value = strTemp
					Else
						rsSectorType.Fields("use2s").Value = vbNullString
					End If
					strTemp = Trim(Mid(strLine, n, 2))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use2n").Value = CShort(strTemp)
					Else
						rsSectorType.Fields("use2n").Value = 0
					End If
				End If
				
				n = InStr(tokenstart, strHeader, "use3")
				If n > 0 Then
					strTemp = Trim(Mid(strLine, n + 3, 1))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use3s").Value = strTemp
					Else
						rsSectorType.Fields("use3s").Value = vbNullString
					End If
					strTemp = Trim(Mid(strLine, n, 2))
					If Len(strTemp) > 0 Then
						rsSectorType.Fields("use3n").Value = CShort(strTemp)
					Else
						rsSectorType.Fields("use3n").Value = 0
					End If
				End If
				
				n = InStr(tokenstart, strHeader, "level")
				If n > 0 Then
					rsSectorType.Fields("level").Value = Trim(Mid(strLine, n, 5))
				End If
				
				n = InStr(tokenstart, strHeader, "min")
				If n > 0 Then
					rsSectorType.Fields("min").Value = CShort(Trim(Mid(strLine, n, 3)))
				End If
				
				n = InStr(tokenstart, strHeader, "lag")
				If n > 0 Then
					rsSectorType.Fields("lag").Value = CShort(Trim(Mid(strLine, n, 3)))
				End If
				
				n = InStr(tokenstart, strHeader, "eff%")
				If n > 0 Then
					rsSectorType.Fields("eff").Value = CShort(Trim(Mid(strLine, n, 4)))
				End If
				
				n = InStr(tokenstart, strHeader, "$$$")
				If n > 0 Then
					rsSectorType.Fields("cost").Value = CShort(Trim(Mid(strLine, n, 4)))
				End If
				
				n = InStr(tokenstart, strHeader, "dep")
				If n > 0 Then
					rsSectorType.Fields("dep").Value = CShort(Trim(Mid(strLine, n, 3)))
				End If
				
				n = InStr(tokenstart, strHeader, " c ")
				If n > 0 Then
					rsSectorType.Fields("pcode").Value = Trim(Mid(strLine, n, 3))
				End If
				If Len(Trim(rsSectorType.Fields("pcode").Value)) > 0 Then
					
					eiItem = Items.FindByLetter(rsSectorType.Fields("pcode").Value)
					'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					If IsReference(eiItem) Then
						eiItem.ProductName = rsSectorType.Fields("product").Value
						eiItem.SaveItem()
					End If
				End If
		End Select
		
		rsSectorType.Update()
		
		Exit Sub
		
		'Error handling routine
ParseShowSector_Exit: 
		EmpireError("ParseShowSector", CStr(n), strLine)
	End Sub
	
	'Parse results of show build string to update database
	Private Sub ParseShowItem(ByRef strLine As String, ByRef Line As Short)
		Dim fldIndex As Short
		Dim n As Short
		'Dim fstart As Integer
		'Dim strSpace As String
		'Dim strDes As String
		'Dim strDesc As String
		'Dim strTemp As String
		Dim strTokens() As String
		Dim chrItem As String
		'Dim strValues As String
		Static strHeader As String
		Static strHeader2 As String
		
		'Command: Show Item
		'Printing for tech level '344'
		'item   value sell lbs   packing   item
		'letter                rg wh ur bk name
		'     c     1   no   1 10 10 10 10 civilians
		'     m     0   no   1  1  1  1  1 military
		'     s     5  yes   1  1 10  1  1 shells
		'     g    60  yes  10  1 10  1  1 guns
		'     p     4  yes   1  1 10  1  1 petrol
		'     i     2  yes   1  1 10  1  1 iron ore
		'     d    20  yes   5  1 10  1  1 dust (gold)
		'     b   280  yes  50  1  5  1  4 bars of gold
		'     f     0  yes   1  1 10  1  1 food
		'     o     8  yes   1  1 10  1  1 oil
		'     l     2  yes   1  1 10  1  1 light products
		'     h     4  yes   1  1 10  1  1 heavy products
		'     u     1  yes   2  1  2  1  1 uncompensated workers
		'     r   150  yes   8  1 10  1  1 radioactive materials
		
		If Line = 2 Then
			strHeader = strLine
			Exit Sub
		End If
		If Line = 3 Then
			strHeader2 = strLine
			Exit Sub
		End If
		
		If Len(Trim(strLine)) = 0 Then
			Exit Sub
		End If
		
		On Error GoTo ParseShowItem_Exit
		
		n = 6
		chrItem = Mid(strLine, 6, 1)
		'n = 2
		'strDesc = Trim$(Mid$(strLine, 3, tokenstart - 3))
		
		rsItems.Seek("=", chrItem)
		If rsItems.NoMatch Then
			rsItems.AddNew()
			'initialize fields
			For fldIndex = 0 To rsItems.Fields.Count - 1
				If rsItems.Fields(fldIndex).Type = DAO.DataTypeEnum.dbText Then
					rsItems.Fields(fldIndex).Value = "???"
				Else
					rsItems.Fields(fldIndex).Value = -1
				End If
			Next fldIndex
			rsItems.Fields("letter").Value = chrItem
		Else
			rsItems.Edit()
		End If
		
		n = InStr(strHeader, "value")
		If n > 0 Then
			rsItems.Fields("value").Value = Val(Trim(Mid(strLine, n, 5)))
		End If
		
		n = InStr(strHeader, "sell")
		If n > 0 Then
			If InStr(Mid(strLine, n, 4), "yes") > 0 Then
				rsItems.Fields("sell").Value = True
			Else
				rsItems.Fields("sell").Value = False
			End If
		End If
		
		n = InStr(strHeader, "lbs")
		If n > 0 Then
			rsItems.Fields("lbs").Value = Val(Trim(Mid(strLine, n, 3)))
		End If
		
		n = InStr(strHeader2, "rg")
		If n > 0 Then
			rsItems.Fields("pack_rg").Value = Val(Trim(Mid(strLine, n, 2)))
		End If
		
		n = InStr(strHeader2, "wh")
		If n > 0 Then
			rsItems.Fields("pack_wh").Value = Val(Trim(Mid(strLine, n, 2)))
		End If
		
		n = InStr(strHeader2, "ur")
		If n > 0 Then
			rsItems.Fields("pack_ur").Value = Val(Trim(Mid(strLine, n, 2)))
		End If
		
		n = InStr(strHeader2, "bk")
		If n > 0 Then
			rsItems.Fields("pack_bk").Value = Val(Trim(Mid(strLine, n, 2)))
		End If
		
		n = InStr(strHeader2, "name")
		If n > 0 Then
			rsItems.Fields("name").Value = Trim(Mid(strLine, n))
		End If
		
		rsItems.Update()
		Exit Sub
		
		'Error handling routine
ParseShowItem_Exit: 
		EmpireError("ParseShowItem", CStr(n), strLine)
	End Sub
	
	Public Sub ParseShowText(ByRef strLine As String, ByRef UnitType As String, ByRef ReportType As String, ByRef LineNumber As Short)
		
		If UnitType = "" Then
			Exit Sub
		End If
		
		'on the first line number - delete what you had
		If LineNumber = 1 And Not (rsShowText.BOF And rsShowText.EOF) Then
			rsShowText.MoveFirst()
			While Not rsShowText.EOF
				If rsShowText.Fields("UnitType").Value = UnitType And rsShowText.Fields("ReportType").Value = ReportType Then
					rsShowText.Delete()
				End If
				rsShowText.MoveNext()
			End While
		End If
		
		rsShowText.AddNew()
		rsShowText.Fields("UnitType").Value = UnitType
		rsShowText.Fields("ReportType").Value = ReportType
		rsShowText.Fields("LineNumber").Value = LineNumber
		rsShowText.Fields("Text").Value = RTrim(strLine)
		rsShowText.Update()
	End Sub
	
	Public Sub ParseShowInfrastructure(ByRef strLine As String, ByRef iLine As Short)
		Dim strID As String
		Dim strTokens() As String
		Dim iIndex As Short
		Dim strParseVariable As String
		
		strParseVariable = "Identifier"
		On Error GoTo ParseShowInfrastructure_Exit
		
		ParseDumpString(strLine, strTokens)
		If strTokens(0) = "Infrastructure" Or strTokens(0) = "type" Or strTokens(0) = "" Or strTokens(0) = "sector" Then
			Exit Sub
		End If
		
		strID = Left(strTokens(0), 5)
		
		'update the build records
		rsBuild.Seek("=", "i", strID)
		If rsBuild.NoMatch Then
			rsBuild.AddNew()
			'initialize fields
			For iIndex = 0 To rsBuild.Fields.Count - 1
				If rsBuild.Fields(iIndex).Type = DAO.DataTypeEnum.dbText Then
					rsBuild.Fields(iIndex).Value = " "
				Else
					rsBuild.Fields(iIndex).Value = 0
				End If
			Next iIndex
			rsBuild.Fields("id").Value = strID
			rsBuild.Fields("type").Value = "i"
		Else
			rsBuild.Edit()
		End If
		If strID = "rail" Or strID = "road" Or strID = "defen" Or (frmOptions.bSPAtlantis And (strID = "radar" Or strID = "fort" Or strID = "runwa")) Then
			strParseVariable = "lcms"
			rsBuild.Fields("lcm").Value = CShort(strTokens(2))
			strParseVariable = "hcms"
			rsBuild.Fields("hcm").Value = CShort(strTokens(3))
			strParseVariable = "mobility"
			rsBuild.Fields("stat2").Value = CShort(strTokens(4))
			strParseVariable = "cost"
			rsBuild.Fields("cost").Value = CShort(strTokens(5))
		ElseIf frmOptions.bSPAtlantis And strID = "navag" Then 
			strParseVariable = "lcms"
			rsBuild.Fields("lcm").Value = CShort(strTokens(1))
			strParseVariable = "hcms"
			rsBuild.Fields("hcm").Value = CShort(strTokens(2))
			strParseVariable = "mobility"
			rsBuild.Fields("stat2").Value = CShort(strTokens(3))
			strParseVariable = "cost"
			rsBuild.Fields("cost").Value = CShort(strTokens(4))
		Else
			strParseVariable = "designation cost"
			rsBuild.Fields("stat1").Value = CShort(strTokens(1))
			strParseVariable = "efficiency cost"
			rsBuild.Fields("cost").Value = CShort(strTokens(2))
			strParseVariable = "lcms"
			rsBuild.Fields("lcm").Value = CShort(strTokens(3))
			strParseVariable = "hcms"
			rsBuild.Fields("hcm").Value = CShort(strTokens(4))
		End If
		
		rsBuild.Update()
		
		Exit Sub 'Error handling routine
		
ParseShowInfrastructure_Exit: 
		EmpireError("ParseShowInfrastructure", strParseVariable, strLine)
	End Sub
	
	Public Function ParseWaypoints(ByRef strWP As String) As Object
		Dim n As Short
		' Dim x As Integer    8/2003 efj  removed
		Dim wp As Object
		
		n = InStr(strWP, ";")
		If n <= 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ParseWaypoints = False
			Exit Function
		End If
		
		ReDim wp(0)
		
		While n > 0
			ReDim Preserve wp(UBound(wp) + 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			wp(UBound(wp)) = Left(strWP, n - 1)
			strWP = Mid(strWP, n + 1)
			n = InStr(strWP, ";")
		End While
		
		'UPGRADE_WARNING: Couldn't resolve default property of object wp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ParseWaypoints = wp
		Exit Function
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseWaypoints", CStr(n), strWP)
	End Function
	
	Public Sub ParseShipView(ByRef strLine As String)
		'[fert:71] fb   fishing boat (#10) @ -2,-4 0% sea
		'[oil:71] oe   oil exploration boat (#12) @ -3,-3 0% sea
		
		Dim n As Short
		Dim sx As Short
		Dim sy As Short
		Dim strSect As String
		On Error GoTo Empire_Error
		
		'We only record sea hexes (not bridges or harbors)
		n = InStr(strLine, " 0% sea")
		If n = 0 Then
			Exit Sub
		End If
		
		'get the x,y location
		strSect = StringBetween(strLine, " @ ", " 0% sea")
		If Not ParseSectors(sx, sy, strSect) Then
			Exit Sub
		End If
		
		'locate or build the database record
		rsSea.Seek("=", sx, sy)
		If rsSea.NoMatch Then
			rsSea.AddNew()
			rsSea.Fields("x").Value = sx
			rsSea.Fields("y").Value = sy
			rsSea.Fields("fert").Value = -1 'negative one mean unknown
			rsSea.Fields("ftimestamp").Value = GetWinACEUTC
			rsSea.Fields("ocontent").Value = -1
			rsSea.Fields("otimestamp").Value = GetWinACEUTC
		Else
			rsSea.Edit()
		End If
		
		'now parse for the fert or the ocontent
		n = InStr(strLine, "[fert:")
		If n > 0 Then
			rsSea.Fields("fert").Value = CShort(StringBetween(strLine, "[fert:", "]"))
			rsSea.Fields("ftimestamp").Value = GetWinACEUTC
		Else
			n = InStr(strLine, "[oil:")
			If n > 0 Then
				rsSea.Fields("ocontent").Value = CShort(StringBetween(strLine, "[oil:", "]"))
			End If
			rsSea.Fields("otimestamp").Value = GetWinACEUTC
		End If
		
		'finally, update the record and exit.
		rsSea.Update()
		If sx = frmDrawMap.SelX And sy = frmDrawMap.SelY Then
			frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
			frmDrawMap.DisplayFirstSectorPanel()
		End If
		Exit Sub
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseShipView", vbNullString, strLine)
	End Sub
	Public Sub ParseAirCombat(ByRef strLine As String)
		' Escher     Savoia      strength int odds  damage           results
		' sf  #401   jes #763     21/12   61  0.36   20/41  cleared       aborted @23%
		' sf  #299   jes #764     14/16   27  0.53   14/13  aborted @32%  cleared
		' sam #735   jes #770     28/16   55  0.36  100/38  shot down     cleared
		' jf2 #403   mb  #8        8/15   39  0.65   21/8   shot down     cleared
		
		Static stFirst As Boolean
		Static stOpponent As Short
		Dim strTemp As String
		Dim n As Short
		Dim strType As String
		Dim planeid As Short
		Dim nDamage As Short
		Dim neff As Short
		On Error GoTo Empire_Error
		
		
		'Find the first plane#
		n = InStr(strLine, "#")
		If n = 0 Then
			'Must be parsing a header
			strTemp = Trim(Left(strLine, 12))
			stOpponent = Nations.NationNumber(strTemp)
			stFirst = True
			'if your the first one, your opponent must be the second!
			If stOpponent = CountryNumber Then
				strTemp = Trim(Mid(strLine, 13, 12))
				stOpponent = Nations.NationNumber(strTemp)
				stFirst = False
			End If
		Else
			neff = -1
			nDamage = -1
			'parsing a plane combat line
			If Not stFirst Then
				n = InStr(n + 1, strLine, "#")
			End If
			strType = Trim(Mid(strLine, n - 4, 4))
			planeid = CShort(Trim(Mid(strLine, n + 1, 5)))
			
			'check for aborted @nn% message
			n = InStr(n + 1, strLine, "@")
			If Not stFirst And n > 0 And n < 70 Then
				n = InStr(n + 1, strLine, "@")
			End If
			
			If n > 0 Then
				neff = CShort(StringBetween(Mid(strLine, n), "@", "%"))
			Else
				'get the damage string
				n = InStr(strLine, "/")
				n = InStr(n + 1, strLine, "/")
				If stFirst Then
					nDamage = CShort(Trim(Mid(strLine, n - 3, 3)))
				Else
					nDamage = CShort(Trim(Mid(strLine, n + 1, 3)))
				End If
			End If
			
			'update the database
			rsAirCombat.Seek("=", stOpponent, planeid)
			If rsAirCombat.NoMatch Then
				rsAirCombat.AddNew()
				rsAirCombat.Fields("nation").Value = stOpponent
				rsAirCombat.Fields("id").Value = planeid
				rsAirCombat.Fields("type").Value = strType
				rsAirCombat.Fields("eff").Value = 100
				rsAirCombat.Fields("numcombats").Value = 1
			Else
				rsAirCombat.Edit()
				rsAirCombat.Fields("numcombats").Value = rsAirCombat.Fields("numcombats").Value + 1
				If rsAirCombat.Fields("nextupdate").Value <> timNextUpdate Then
					rsAirCombat.Fields("eff").Value = 100
					rsAirCombat.Fields("type").Value = strType
					rsAirCombat.Fields("numcombats").Value = 1
				End If
			End If
			
			rsAirCombat.Fields("nextupdate").Value = timNextUpdate
			
			If neff >= 0 Then
				rsAirCombat.Fields("eff").Value = neff
			Else
				If nDamage > 0 Then
					rsAirCombat.Fields("eff").Value = rsAirCombat.Fields("eff").Value - nDamage
					If rsAirCombat.Fields("eff").Value < 10 Then
						rsAirCombat.Fields("eff").Value = 0
					End If
				End If
			End If
			'set the plane status
			If rsAirCombat.Fields("eff").Value >= 40 Then
				rsAirCombat.Fields("status").Value = "A"
			ElseIf rsAirCombat.Fields("eff").Value >= 10 Then 
				rsAirCombat.Fields("status").Value = "D"
			Else
				rsAirCombat.Fields("status").Value = "S"
			End If
			
			'finally, update the record and exit.
			rsAirCombat.Update()
		End If
		
		Exit Sub
		'Error handling routine
Empire_Error: 
		EmpireError("ParseAirCombat", vbNullString, strLine)
	End Sub
	
	Public Function GetUnitClass(ByRef strID As String, ByRef strLine As String) As String
		Dim bPlane As Boolean
		Dim bLand As Boolean
		Dim bShip As Boolean
		Dim sdesc As String
		Dim ldesc As String
		Dim pdesc As String
		On Error GoTo Empire_Error
		
		'make sure we have a valid string
		If Len(strID) = 0 Then
			GetUnitClass = vbNullString
			Exit Function
		End If
		
		'check land
		rsBuild.Seek("=", "l", strID)
		bLand = Not rsBuild.NoMatch
		If bLand Then
			ldesc = rsBuild.Fields("desc").Value
		End If
		
		'check plane
		rsBuild.Seek("=", "p", strID)
		bPlane = Not rsBuild.NoMatch
		If bPlane Then
			pdesc = rsBuild.Fields("desc").Value
		End If
		
		'check ship
		rsBuild.Seek("=", "s", strID)
		bShip = Not rsBuild.NoMatch
		If bShip Then
			sdesc = rsBuild.Fields("desc").Value
		End If
		
		'sub is special - from sonar
		If strID = "sub" Then
			bShip = True
			sdesc = "sub"
		End If
		
		'if not found, then Add and Exit
		If Not (bShip Or bLand Or bPlane) Then
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmAddUnitDes)
			frmAddUnitDes.Text = "Unit Designation " & strID & " not found in Database"
			frmAddUnitDes.Label1.Text = strLine
			frmAddUnitDes.cmdCancel.Text = "Ignore"
			frmAddUnitDes.txtID.Text = strID
			frmAddUnitDes.txtDesc.Text = StringBetween(strLine, ")", "(")
			frmAddUnitDes.txtID.ReadOnly = True
			frmAddUnitDes.ShowDialog()
			GetUnitClass = frmAddUnitDes.UnitClass
			If Len(GetUnitClass) > 0 Then
				rsBuild.AddNew()
				rsBuild.Fields("type").Value = frmAddUnitDes.UnitClass
				rsBuild.Fields("id").Value = strID
				rsBuild.Fields("desc").Value = frmAddUnitDes.txtDesc.Text
				rsBuild.Update()
			End If
			frmAddUnitDes.Close()
			Exit Function
		End If
		
		'Determine if we have a ship, land, or plane
		If bShip Then
			GetUnitClass = "s"
			If bPlane Then
				If InStr(strLine, sdesc) = 0 Then
					If InStr(strLine, pdesc) > 0 Then
						GetUnitClass = "p"
					End If
				End If
			End If
			If bLand Then
				If InStr(strLine, sdesc) = 0 Then
					If InStr(strLine, ldesc) > 0 Then
						GetUnitClass = "l"
					End If
				End If
			End If
		ElseIf bLand Then 
			GetUnitClass = "l"
			If bPlane Then
				If InStr(strLine, ldesc) = 0 Then
					If InStr(strLine, pdesc) > 0 Then
						GetUnitClass = "p"
					End If
				End If
			End If
		Else
			GetUnitClass = "p"
		End If
		
		Exit Function
		'Error handling routine
Empire_Error: 
		EmpireError("GetUnitClass", strID, strLine)
		
	End Function
	
	Public Function ParseShellDamage(ByRef strLine As String) As Short
		On Error GoTo Empire_Error
		'Shell hit sector -3,-1 for 40 damage.
		ParseShellDamage = CShort(Trim(StringBetween(strLine, "for", "damage")))
		'Error handling routine
		Exit Function
Empire_Error: 
		EmpireError("ParseShellDamage", vbNullString, strLine)
	End Function
	
	
	'Parse the results of a sorder command
	Public Sub ParseShipOrders(ByRef strLine As String, Optional ByRef Deityflag As Boolean = False)
		
		On Error GoTo Empire_Error
		
		'Deity version
		'own shp#     ship type      x,y    start   end   len  eta
		'  2    2 dd   destroyer    15,-1   17,-7          6    1
		'  2    3 dd   destroyer    -4,6     1,1           5    1
		'  2   13 cs   cargo ship   -4,6    -5,5   -3,13  12    1
		
		'Player version
		'shp#     ship type      x,y    start   end   len  eta
		'   1 tt   troop trans  -8,-16   0,0        no route possible
		'   2 dd   destroyer    14,-10  16,-16         6    1
		'   3 dd   destroyer    -5,-3    0,-8          5    1
		'  13 cs   cargo ship   -5,-3   -6,-4  -4,4   12    1
		'  16 ft   fishing tra  -5,-3   -5,-3  -6,-6  suspended
		'  17 ft   fishing tra  -5,-3   -5,-3  -6,-6  suspended
		'  26 fb   fishing boa  11,-1   11,-1  13,-1  loading
		'   3 cs   cargo ship   -4,-2   -3,-1 -21,-3  11    1
		'       test 1 2 3
		'   4 fb   fishing boa   0,0   has a sail path
		'
		
		Dim n As Short
		Dim id As Short
		Dim owner As String
		Dim strTokens() As String
		' Dim numTokens As Integer    8/2003 efj  removed
		
		'if this is a deity dump, include special handling for country code
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Deityflag) Then
			Deityflag = False
		End If
		
		'110603 rjk: sorder on a sector gives an error
		'200204 rjk: Added check for ship name in order output, find using the blanks at the beginning
		If InStr(strLine, "#") = 0 And InStr(strLine, ",") > 0 And Left(strLine, 7) <> "       " And InStr(strLine, "No ship") = 0 Then
			ParseDumpString(strLine, strTokens)
			
			n = LBound(strTokens)
			If Deityflag Then
				owner = strTokens(n)
				n = n + 1
			Else
				owner = "0"
			End If
			
			id = CShort(strTokens(n))
			rsShipOrders.Seek("=", id)
			If Not rsShipOrders.NoMatch Then
				rsShipOrders.Edit()
			Else
				rsShipOrders.AddNew()
				rsShipOrders.Fields("id").Value = id
			End If
			rsShipOrders.Fields("owner").Value = CShort(owner)
			
			'skip ship type and ship name
			While InStr(strTokens(n), ",") = 0
				n = n + 1
			End While
			
			rsShipOrders.Fields("curr sector").Value = strTokens(n)
			n = n + 1
			
			'drk@unxsoft.com 7/20/02.  if a ship has "sail" orders, it still shows up
			'here, but the rest of the line is blank
			If (UBound(strTokens, 1) < n) Or (InStr(strLine, "has a sail path") > 0) Then
				rsShipOrders.Fields("start sector").Value = "Sail"
			Else
				rsShipOrders.Fields("start sector").Value = strTokens(n)
				
				n = n + 1
				
				If (strTokens(n) = "no") Then ' "no route possible"
					'drk@unxsoft.com 7/14/02.  if we f*cked up the order command, e.g. giving a
					'land sector as a dest, then we'll come here.  This might not be the best
					'way to handle this situation, but at least we won't crash on a parse error
					'like before.
					rsShipOrders.Fields("length").Value = 999
					rsShipOrders.Fields("eta").Value = 999
					rsShipOrders.Fields("end sector").Value = "???"
				Else
					If InStr(strTokens(n), ",") > 0 Then
						If InStr(strTokens(n), "has") > 0 Then '102603 rjk: Server bug no space between End Sector and 'has arrived'.
							rsShipOrders.Fields("end sector").Value = Left(strTokens(n), InStr(strTokens(n), "has") - 1)
							strTokens(n) = Mid(strTokens(n), InStr(strTokens(n), "has"))
						ElseIf InStr(strTokens(n), "no") > 0 Then  '102603 rjk: Server bug no space between End Sector and 'no route possible'.
							rsShipOrders.Fields("end sector").Value = Left(strTokens(n), InStr(strTokens(n), "no") - 1)
							strTokens(n) = Mid(strTokens(n), InStr(strTokens(n), "no"))
						ElseIf InStr(strTokens(n), "route") > 0 Then  '102603 rjk: Server bug no space between End Sector and 'route too long'.
							rsShipOrders.Fields("end sector").Value = Left(strTokens(n), InStr(strTokens(n), "route") - 1)
							strTokens(n) = Mid(strTokens(n), InStr(strTokens(n), "route"))
						Else
							rsShipOrders.Fields("end sector").Value = strTokens(n)
							n = n + 1
						End If
					End If
					
					'  suspended                      loading
					If Left(strTokens(n), 1) = "s" Or Left(strTokens(n), 1) = "l" Then
						rsShipOrders.Fields("length").Value = 999
						rsShipOrders.Fields("eta").Value = 999
					ElseIf (strTokens(n) = "has") Then  '"has arrived".  why they don't just say eta = 0, I don't know... drk@unxsoft.com 10/31/02
						rsShipOrders.Fields("eta").Value = 0
						rsShipOrders.Fields("length").Value = 999
					ElseIf (strTokens(n) = "no") Then  '"no route possible".  102603 rjk: Added based on the server code
						rsShipOrders.Fields("eta").Value = 0
						rsShipOrders.Fields("length").Value = 999
					ElseIf (strTokens(n) = "route") Then  '"route too long".  102603 rjk: Added based on the server code
						rsShipOrders.Fields("eta").Value = 0
						rsShipOrders.Fields("length").Value = 999
					Else
						rsShipOrders.Fields("length").Value = strTokens(n)
						rsShipOrders.Fields("eta").Value = strTokens(n + 1)
					End If
				End If
			End If
			
			rsShipOrders.Update()
		End If
		
		If frmDrawMap.PromptUp And frmDrawMap.PromptForm Is frmPromptOrder Then
			frmPromptOrder.OrderLoadData()
		End If
		Exit Sub
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseShipOrders", CStr(n), strLine)
	End Sub
	
	
	'Parse the results of a qorder command
	Public Sub ParseCargoOrders(ByRef strLine As String, Optional ByRef Deityflag As Boolean = False)
		
		'shp#     ship type    [Starting]       (Ending)
		'   2 dd   destroyer   [] , ()
		'   3 dd   destroyer   [] , ()
		'  13 cs   cargo ship  [1-l:100 2-s:1 3-h:50 ] , (1-l:1 2-s:50 3-h:1 )
		'  14 ft   fishing tra [1-f:15 ] , (1-f:15 )
		'  15 ft   fishing tra [1-f:15 ] , (1-f:15 )
		'  16 ft   fishing tra [1-f:15 ] , (1-f:15 )
		'  17 ft   fishing tra [1-f:15 ] , (1-f:15 )
		
		Dim n As Short
		Dim id As Short
		Dim owner As String
		Dim strTokens() As String
		Dim numTokens As Short
		Dim strTemp As String
		Dim strHold As String
		Dim strItem As String
		Dim strLevel As String
		
		
		'if this is a deity dump, include special handling for country code
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Deityflag) Then
			Deityflag = False
		End If
		
		On Error GoTo Empire_Error
		
		If InStr(strLine, "#") = 0 And InStr(strLine, "[") > 0 Then
			
			strTemp = Left(strLine, InStr(strLine, "[") - 1)
			ParseDumpString(strTemp, strTokens)
			numTokens = UBound(strTokens)
			
			n = LBound(strTokens)
			If Deityflag Then
				owner = strTokens(n)
				n = n + 1
			Else
				owner = "0"
			End If
			
			id = CShort(strTokens(n))
			rsShipOrders.Seek("=", id)
			If Not rsShipOrders.NoMatch Then
				rsShipOrders.Edit()
			Else
				rsShipOrders.AddNew()
				rsShipOrders.Fields("id").Value = id
			End If
			rsShipOrders.Fields("owner").Value = CShort(owner)
			
			'clear current cargo settings
			For n = 1 To 6
				strHold = CStr(n)
				rsShipOrders.Fields("start cargo " & strHold).Value = vbNullString
				rsShipOrders.Fields("end cargo " & strHold).Value = vbNullString
				rsShipOrders.Fields("start level " & strHold).Value = 0
				rsShipOrders.Fields("end level " & strHold).Value = 0
			Next n
			
			'get cargo set one
			strTemp = StringBetween(strLine, "[", "]")
			If Len(strTemp) > 0 Then
				ParseDumpString(strTemp, strTokens)
				numTokens = UBound(strTokens)
				For n = LBound(strTokens) To numTokens
					strHold = StringBetween(strTokens(n), vbNullString, "-")
					strItem = StringBetween(strTokens(n), "-", ":")
					strLevel = CStr(CShort(StringBetween(strTokens(n), ":")))
					
					rsShipOrders.Fields("start cargo " & strHold).Value = strItem
					rsShipOrders.Fields("start level " & strHold).Value = strLevel
				Next n
			End If
			
			'get cargo set two
			strTemp = StringBetween(strLine, "(", ")")
			If Len(strTemp) > 0 Then
				ParseDumpString(strTemp, strTokens)
				numTokens = UBound(strTokens)
				For n = LBound(strTokens) To numTokens
					strHold = StringBetween(strTokens(n), vbNullString, "-")
					strItem = StringBetween(strTokens(n), "-", ":")
					strLevel = CStr(CShort(StringBetween(strTokens(n), ":")))
					
					rsShipOrders.Fields("end cargo " & strHold).Value = strItem
					rsShipOrders.Fields("end level " & strHold).Value = strLevel
				Next n
			End If
			
			rsShipOrders.Update()
		End If
		
		If frmDrawMap.PromptUp And frmDrawMap.PromptForm Is frmPromptOrder Then
			frmPromptOrder.OrderLoadData()
		End If
		Exit Sub
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseCargoOrders", vbNullString, strLine)
	End Sub
	
	'Parse the results of a sorder command
	Public Sub ParseMissions(ByRef strLine As String, ByRef strType As String, Optional ByRef Deityflag As Boolean = False)
		
		On Error GoTo Empire_Error
		
		'Player version
		'land
		'inf  infantry #5 on a reserve mission with maximum reaction radius 4
		'hat  hvy artillery #8 on an interdiction mission, centered on 5,-3, radius 7'
		
		'plane
		'jf2  AV-8B Harrier #4 on an air defense mission, centered on -2,-2, radius 5
		'jl   A-6 Intruder #26 on an interdiction mission, centered on -2,-2, radius 3
		'ssm  V1 #29 on an interdiction mission, centered on -2,-2, radius 6
		'jes  F-14E jet escort #31 on an escort mission
		'jes  F-14E jet escort #32 on an escort mission
		'jf2  AV-8B Harrier #3 on a support mission, centered on -2,-2, radius 5
		'jf2  AV-8B Harrier #0 on a offensive support mission, centered on -2,-2, radius 11
		'jf2  AV-8B Harrier #1 on a defensive support mission, centered on -2,-2, radius 11
		
		'Ship
		'ncr  nuc cruiser (#8) on an interdiction mission, centered on -5,-3, radius 7
		'
		'Command: mission Ship * q - 5, -3
		'Thing                         x,y   op-sect rad mission
		'dd   destroyer (#0)        -10,-8               has no mission.
		'ncr  nuc cruiser (#6)       -5,-3   -5,-5     1 is on an interdiction mission
		'ncr  nuc cruiser (#8)       -5,-3   -5,-3     7 is on an interdiction mission
		'ncr  nuc cruiser (#9)       -5,-3               has no mission.
		'ft   fishing trawler (#14)   -5,-3               has no mission.
		
		'Command: mission l * q - 5, -3
		'Thing                         x,y   op-sect rad mission
		'inf  infantry #0             3,-1    3,-1     3 is on a reserve mission
		'inf  infantry #1             5,-3               has no mission.
		'inf  infantry #5             5,-3    5,-3     0 is on a reserve mission
		'hat  hvy artillery #6       -6,-4               has no mission.
		'hat  hvy artillery #7        5,-3    5,-3     3 is on an interdiction mission
		'hat  hvy artillery #8        5,-3    5,-3     7 is on an interdiction mission
		'eng  engineer #12            5,-3               has no mission.
		
		
		
		'Command: mission p * q - 5, -3
		'Thing                         x,y   op-sect rad mission
		'jf2  AV-8B Harrier #0       -2,-2   -2,-2    11 is on a offensive support mission
		'jf2  AV-8B Harrier #1       -2,-2   -2,-2    11 is on a defensive support mission
		'jf2  AV-8B Harrier #2       -2,-2               has no mission.
		'jf2  AV-8B Harrier #3       -2,-2   -2,-2     5 is on a support mission
		'jf2  AV-8B Harrier #5       -2,-2   -2,-2     5 is on an air defense mission
		'jf2  AV-8B Harrier #6       -2,-2               has no mission.
		'jl   A-6 Intruder #26       -2,-2   -2,-2     3 is on an interdiction mission
		'jl   A-6 Intruder #27       -2,-2   -2,-2     3 is on an interdiction mission
		'ssm  V1 #28                 -2,-2               has no mission.
		'ssm  V1 #30                 -2,-2   -2,-2     6 is on an interdiction mission
		'jes  F-14E jet escort #31   -2,-2            16 is on an escort mission
		'jes  F-14E jet escort #32   -2,-2            16 is on an escort mission
		'jes  F-14E jet escort #33   -2,-2               has no mission.
		
		Dim n As Short
		Dim n1 As Short
		Dim id As Short
		Dim strTokens() As String
		Dim strTemp As String
		Dim var As String
		Dim bFillGrid As Boolean
		
		bFillGrid = False
		
		'if this is a deity dump, include special handling for country code
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(Deityflag) Then
			Deityflag = False
		End If
		
		On Error GoTo Empire_Error
		
		If InStr(strLine, "(") Then
			strTemp = ")"
		Else
			strTemp = " "
		End If
		
		var = "id"
		If InStr(strLine, "has no mission") > 0 Then
			'200204 rjk: Added check for # in the ship name
			If strTemp = ")" Then
				id = CShort(StringBetween(strLine, "(#", strTemp))
			Else
				id = CShort(StringBetween(strLine, "#", strTemp))
			End If
			rsMissions.Seek("=", strType, id)
			If Not rsMissions.NoMatch Then
				rsMissions.Delete()
			End If
			n = InStr(strLine, ",")
			If n > 0 Then
				n1 = n
				While Mid(strLine, n, 1) <> " "
					n = n - 1
				End While
				While Mid(strLine, n1, 1) <> " "
					n1 = n1 + 1
				End While
				If Mid(strLine, n + 1, n1 - n - 1) = CStr(frmDrawMap.SelX) & "," & CStr(frmDrawMap.SelY) Then
					bFillGrid = True
				End If
			End If
		ElseIf InStr(strLine, "on a") > 0 Then 
			'200204 rjk: Added check for # in the ship name
			If strTemp = ")" Then
				id = CShort(StringBetween(strLine, "(#", strTemp))
			Else
				id = CShort(StringBetween(strLine, "#", strTemp))
			End If
			rsMissions.Seek("=", strType, id)
			If Not rsMissions.NoMatch Then
				rsMissions.Edit()
			Else
				rsMissions.AddNew()
				rsMissions.Fields("id").Value = id
				rsMissions.Fields("type").Value = strType
			End If
			var = "owner"
			rsMissions.Fields(var).Value = 0
			
			'check if from a query
			If InStr(strLine, "is on") > 0 Then
				n = InStr(strLine, ",")
				While Mid(strLine, n, 1) <> " "
					n = n - 1
				End While
				ParseDumpString(Mid(strLine, n), strTokens)
				'        numTokens = UBound(strTokens)    8/2003 efj  removed
				If strTokens(0) = CStr(frmDrawMap.SelX) & "," & CStr(frmDrawMap.SelY) Then
					bFillGrid = True
				End If
				n = LBound(strTokens) + 1
				var = "op sector"
				If InStr(strTokens(n), ",") > 0 Then
					rsMissions.Fields(var).Value = strTokens(n)
					n = n + 1
				Else
					rsMissions.Fields(var).Value = vbNullString
				End If
				var = "radius"
				rsMissions.Fields(var).Value = strTokens(n)
				var = "mission"
				strTemp = StringBetween(strLine, "is on a", "mission")
				If Left(strTemp, 2) = "n " Then
					strTemp = Mid(strTemp, 3)
				End If
				rsMissions.Fields(var).Value = strTemp
			Else
				var = "op sector"
				If InStr(strLine, "centered") > 0 Then
					rsMissions.Fields(var).Value = StringBetween(strLine, "centered on", ", ")
				Else
					rsMissions.Fields(var).Value = vbNullString
				End If
				
				var = "radius"
				If InStr(strLine, var) > 0 Then
					rsMissions.Fields(var).Value = CShort(StringBetween(strLine, "radius"))
				Else
					rsMissions.Fields(var).Value = 0
				End If
				var = "mission"
				strTemp = StringBetween(strLine, " on a", "mission")
				If Left(strTemp, 2) = "n " Then
					strTemp = Mid(strTemp, 3)
				End If
				rsMissions.Fields(var).Value = strTemp
			End If
			n = n + 1
			
			rsMissions.Update()
		End If
		
		If bFillGrid Then
			frmDrawMap.FillGrid()
		End If
		
		Exit Sub
		
		'Error handling routine
Empire_Error: 
		EmpireError("ParseMissions", var, strLine)
	End Sub
	
	Public Sub ParseAndDeleteUnits(ByRef strResponse As String)
		'linf light infantry #5 blown up by the crew!
		Dim strUnit As String
		Dim strID As String
		Dim iInteger As Short
		Dim iPos As Short
		Dim iPos2 As Short
		Dim hldIndex As String
		
		hldIndex = rsEnemyUnit.Index
		On Error GoTo Empire_Error
		
		rsEnemyUnit.Index = "ByID"
		
		strUnit = Left(strResponse, 4)
		
		iPos = InStr(strResponse, "#")
		If iPos > 0 Then
			iPos2 = InStr(iPos, strResponse, " ")
			If iPos2 > 0 Then
				strID = Mid(strResponse, iPos + 1, iPos2 - iPos - 1)
				'check land
				rsBuild.Seek("=", "l", strUnit)
				If Not rsBuild.NoMatch Then
					rsEnemyUnit.Seek("=", "l", strID)
					If Not rsEnemyUnit.NoMatch Then
						rsEnemyUnit.Delete()
					End If
				Else
					'check plane
					rsBuild.Seek("=", "p", strUnit)
					If Not rsBuild.NoMatch Then
						rsEnemyUnit.Seek("=", "p", strID)
						If Not rsEnemyUnit.NoMatch Then
							rsEnemyUnit.Delete()
						End If
					End If
				End If
			End If
		End If
		rsEnemyUnit.Index = hldIndex
		Exit Sub
		
Empire_Error: 
		rsEnemyUnit.Index = hldIndex
		EmpireError("ParseAndDeleteUnits", strResponse, "")
	End Sub
End Module