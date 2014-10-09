VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'050206 rjk: Created
'220506 rjk: Changed NSC_STRINGY constant to variable extracted from meta-type table.
'150606 Convert iNscSTRINGY to Symbol table.

Private mCol As Collection
Private mstrName As String
Private mxdType As enumXdumpType

Public Property Get Column(vntIndexKey As Variant) As EmpTableColumn
  'used when referencing an element in the collection
  'vntIndexKey contains either the Index or Key to the collection,
  'this is why it is declared as a Variant
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
Set Column = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
Count = mCol.Count
End Property

Public Property Get ParameterCount() As Long
Dim Item As EmpTableColumn

ParameterCount = 0
For Each Item In mCol
    If estsTable(XD_META_TYPE).Count >= Item.ColType Then
        If estsTable(XD_META_TYPE)(Item.ColType).Name <> "c" And Item.Length > 1 Then
            ParameterCount = ParameterCount + Item.Length
        Else
            ParameterCount = ParameterCount + 1
        End If
    Else
        ParameterCount = ParameterCount + 1
    End If
Next Item
End Property

Public Property Get NewEnum() As IUnknown
'this property allows you to enumerate
'this collection with the For...Each syntax
Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Let Name(ByVal vData As String)
mstrName = vData
End Property

Public Property Get Name() As String
Name = mstrName
End Property

Public Property Let TblType(ByVal vData As enumXdumpType)
mxdType = vData
End Property

Public Property Get TblType() As enumXdumpType
TblType = mxdType
End Property

Private Sub Class_Initialize()
'creates the collection when this class is created
Set mCol = New Collection

End Sub

Public Sub AddUpdateColumn(strLine As String)
Dim etTableColumn As EmpTableColumn
Dim strParams() As String
Dim strName As String

strParams = Split(strLine, " ")
strName = RemoveEscapeSequencesAndQuotes(strParams(0))
On Error GoTo Add_New
Set etTableColumn = mCol(strName)
etTableColumn.ColType = CInt(strParams(1))
etTableColumn.Flags = CInt(strParams(2))
etTableColumn.Length = CInt(strParams(3))
etTableColumn.Table = CInt(strParams(4))
Exit Sub
Add_New:
Set etTableColumn = New EmpTableColumn
etTableColumn.Name = strName
etTableColumn.ColType = CInt(strParams(1))
etTableColumn.Flags = CInt(strParams(2))
etTableColumn.Length = CInt(strParams(3))
etTableColumn.Table = CInt(strParams(4))
mCol.Add etTableColumn, strName
End Sub

Public Function GetZeroIndex(strName As String) As Integer
Dim Item As EmpTableColumn
Dim i As Integer

i = 0
For Each Item In mCol
    If Item.Name = strName Then
        GetZeroIndex = i
        Exit Function
    End If
    i = i + 1
Next Item
GetZeroIndex = -1
End Function

Public Function GetOneIndex(strName As String) As Integer
GetOneIndex = GetZeroIndex(strName)
If GetOneIndex <> -1 Then
    GetOneIndex = GetOneIndex + 1
End If
End Function
