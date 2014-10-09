Option Strict Off
Option Explicit On
Friend Class EmpSectorMarker
	Dim sm As New Collection
	Const SECTOR_MARKER_DURATION As Short = 5 'default duration in minutes
	
	
	Public Sub Add(ByRef sx As Short, ByRef sy As Short, ByRef Nation As Short, ByRef Message As String)
		
		Dim col As Collection
		Dim ncol As New Collection
		
		'first look for it
		For	Each col In sm
			'UPGRADE_WARNING: Couldn't resolve default property of object col.item(y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object col.item(x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If col.Item("x") = sx And col.Item("y") = sy Then
				col.Remove("msg")
				col.Remove("ts")
				col.Remove("nation")
				col.Add(Message, "msg")
				col.Add(Now, "ts")
				col.Add(Nation, "nation")
				Exit Sub
			End If
		Next col
		
		ncol.Add(sx, "x")
		ncol.Add(sy, "y")
		ncol.Add(Message, "msg")
		ncol.Add(Nation, "nation")
		ncol.Add(Now, "ts")
		
		sm.Add(ncol)
		
	End Sub
	
	Public Function Find(ByRef sx As Short, ByRef sy As Short) As String
		
		Dim col As Collection
		
		'first look for it
		For	Each col In sm
			'UPGRADE_WARNING: Couldn't resolve default property of object col.item(y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object col.item(x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If col.Item("x") = sx And col.Item("y") = sy Then
				'UPGRADE_WARNING: Couldn't resolve default property of object col.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Find = col.Item("msg")
				Exit Function
			End If
		Next col
		
		Find = vbNullString
		
	End Function
	
	Public Sub Clear()
		
		While Me.Count > 0
			Me.Remove(1)
		End While
		
	End Sub
	
	Public Function Count() As Short
		Count = sm.Count()
	End Function
	
	Public Function X(ByRef Index As Short) As Short
		If Index < 1 Or Index > sm.Count() Then
			X = 0
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sm.item().item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = sm.Item(Index).item("x")
	End Function
	Public Function Y(ByRef Index As Short) As Short
		If Index < 1 Or Index > sm.Count() Then
			Y = 0
			Exit Function
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object sm.item().item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Y = sm.Item(Index).item("y")
	End Function
	'Public Function Message(Index As Integer) As String    8/2003 efj  removed
	'If Index < 1 Or Index > sm.Count Then
	'    Message = vbnullstring
	'    Exit Function
	'End If
	'Message = sm.Item(Index).Item("msg")
	'End Function
	Public Function Nation(ByRef Index As Short) As Short
		If Index < 1 Or Index > sm.Count() Then
			Nation = 0
			Exit Function
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object sm.item().item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Nation = sm.Item(Index).item("nation")
	End Function
	
	Public Sub Remove(ByRef Index As Short)
		Dim col As Collection
		col = sm.Item(Index)
		col.Remove("x")
		col.Remove("y")
		col.Remove("msg")
		col.Remove("ts")
		sm.Remove(Index)
	End Sub
	Public Sub RemoveExpired()
		
		Dim col As Collection
		Dim n As Short
		For n = sm.Count() To 1 Step -1
			col = sm.Item(n)
			'UPGRADE_WARNING: Couldn't resolve default property of object col.item(ts). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Now > DateAdd(Microsoft.VisualBasic.DateInterval.Minute, CDbl(SECTOR_MARKER_DURATION), col.Item("ts")) Then
				Me.Remove(n)
			End If
		Next n
		
	End Sub
End Class