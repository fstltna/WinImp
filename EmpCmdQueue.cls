VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpCmdQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim Queue As Collection
Dim CmdList() As String


'Adds a new command to the queue
Public Sub AddCommand(strCmd As String)
Queue.Add strCmd
End Sub

'Gets the first command from the queue
Public Function GetCommand() As String
If Queue.Count = 0 Then
    GetCommand = vbNullString
    Exit Function
End If

GetCommand = Queue.item(1)
Queue.Remove 1
End Function

'Returns the number of commands pending
Public Function CmdsPending() As Integer
CmdsPending = Queue.Count
End Function

'Clears the queue
Public Sub ClearAllCmds()
While Queue.Count > 0
    Queue.Remove 1
Wend
End Sub

Private Sub Class_Initialize()
Set Queue = New Collection
End Sub

Private Sub Class_Terminate()
ClearAllCmds
Set Queue = Nothing
End Sub

Public Function GetPreviousCommand(clstIndex As Integer) As String

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

Public Function PreviousCommandCount() As Integer
On Error Resume Next
PreviousCommandCount = 0
PreviousCommandCount = UBound(CmdList)

End Function

Public Sub AddtoHistory(strCmd As String)
    
ReDim Preserve CmdList(PreviousCommandCount + 1)
CmdList(PreviousCommandCount) = strCmd

End Sub

'Check to see if a batch query response is next
Public Function QueryResponseNext() As Boolean
If Queue.Count = 0 Then
    QueryResponseNext = False
    Exit Function
End If
QueryResponseNext = (Left$(Queue.item(1), 1) = "&")
End Function


