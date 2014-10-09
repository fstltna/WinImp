Option Strict Off
Option Explicit On
Friend Class EmpCmdQueue
	Dim Queue As Collection
	Dim CmdList() As String
	
	
	'Adds a new command to the queue
	Public Sub AddCommand(ByRef strCmd As String)
		Queue.Add(strCmd)
	End Sub
	
	'Gets the first command from the queue
	Public Function GetCommand() As String
		If Queue.Count() = 0 Then
			GetCommand = vbNullString
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Queue.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCommand = Queue.Item(1)
		Queue.Remove(1)
	End Function
	
	'Returns the number of commands pending
	Public Function CmdsPending() As Short
		CmdsPending = Queue.Count()
	End Function
	
	'Clears the queue
	Public Sub ClearAllCmds()
		While Queue.Count() > 0
			Queue.Remove(1)
		End While
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Queue = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		ClearAllCmds()
		'UPGRADE_NOTE: Object Queue may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Queue = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function GetPreviousCommand(ByRef clstIndex As Short) As String
		
		'If no commands have been entered, an error is generated
		On Error GoTo GetPreviousCommand_Error
		
		If clstIndex <= PreviousCommandCount Then
			GetPreviousCommand = CmdList(clstIndex)
		Else
			GetPreviousCommand = vbNullString
		End If
		
		Exit Function
		
GetPreviousCommand_Error: 
		GetPreviousCommand = vbNullString
		
	End Function
	
	Public Function PreviousCommandCount() As Short
		On Error Resume Next
		PreviousCommandCount = 0
		PreviousCommandCount = UBound(CmdList)
		
	End Function
	
	Public Sub AddtoHistory(ByRef strCmd As String)
		
		ReDim Preserve CmdList(PreviousCommandCount + 1)
		CmdList(PreviousCommandCount) = strCmd
		
	End Sub
	
	'Check to see if a batch query response is next
	Public Function QueryResponseNext() As Boolean
		If Queue.Count() = 0 Then
			QueryResponseNext = False
			Exit Function
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Queue.item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		QueryResponseNext = (Left(Queue.Item(1), 1) = "&")
	End Function
End Class