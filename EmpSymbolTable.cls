VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpSymbolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'110606 rjk: Created
'190606 rjk: Convert value to long to support ship-chr-flags.

Private mCol As Collection
Private mstrName As String
Private mxdType As enumXdumpType

Public Property Get Entry(lValue As Long) As EmpSymbolEntry
Attribute Entry.VB_UserMemId = 0
Set Entry = mCol(CStr(lValue))
End Property

Public Property Get Count() As Long
Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Let TblType(ByVal vData As enumXdumpType)
mxdType = vData
End Property

Public Property Get TblType() As enumXdumpType
TblType = mxdType
End Property

Public Property Let Name(ByVal vData As String)
mstrName = vData
End Property

Public Property Get Name() As String
Name = mstrName
End Property

Private Sub Class_Initialize()
'creates the collection when this class is created
Set mCol = New Collection
End Sub

Public Sub AddUpdateSymbolEntry(lValue As Long, strName As String)
Dim esSymbolEntry As EmpSymbolEntry

On Error GoTo Add_New
Set esSymbolEntry = mCol(CStr(lValue))
esSymbolEntry.Name = strName
Exit Sub
Add_New:
Set esSymbolEntry = New EmpSymbolEntry
esSymbolEntry.Name = strName
esSymbolEntry.Value = lValue
mCol.Add esSymbolEntry, CStr(lValue)
End Sub

