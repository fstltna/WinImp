Attribute VB_Name = "UTF8"
Option Explicit

'Change Log:
'220304 rjk: Created to support UTF8 communication with the server
'301204 rjk: Added to remove SO/SI characters from the server used to indicate
'            highlighting in UTF8 mode.
'260505 rjk: Switched UTF8 from an registry option to login option.

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public Const CP_UTF8 = 65001

'Purpose:Convert Utf8 to Unicode
Public Function UTF8_Decode(baUTF8() As Byte) As String
Dim lUtf8Size As Long
Dim strBuffer As String
Dim lBufferSize As Long
Dim lResult As Long
Dim lPos As Long

UTF8_Decode = vbNullString

lUtf8Size = UBound(baUTF8) + 1
If lUtf8Size <= 0 Then Exit Function

'Set buffer for longest possible string i.e. each byte is
'ANSI, thus 1 unicode(2 bytes)for every utf-8 character.
lBufferSize = lUtf8Size * 2
strBuffer = String$(lBufferSize, vbNullChar)
'Translate using code page 65001(UTF-8)
lResult = MultiByteToWideChar(CP_UTF8, 0, baUTF8(0), lUtf8Size, StrPtr(strBuffer), lBufferSize)
strBuffer = Left$(strBuffer, lResult)
'301204 rjk: Added to remove SO/SI characters from the server used to indicate
'            highlighting in UTF8 mode.
lPos = 1
While InStr(lPos, strBuffer, Chr$(14)) > 0
    lPos = InStr(strBuffer, Chr$(14))
    strBuffer = Left$(strBuffer, lPos - 1) + Right$(strBuffer, Len(strBuffer) - lPos)
Wend
lPos = 1
While InStr(lPos, strBuffer, Chr$(15)) > 0
    lPos = InStr(strBuffer, Chr$(15))
    strBuffer = Left$(strBuffer, lPos - 1) + Right$(strBuffer, Len(strBuffer) - lPos)
Wend
'Trim result to actual length
If lResult Then
   UTF8_Decode = strBuffer
End If
End Function

'Purpose:Convert Unicode string to UTF-8.
Public Function UTF8_Encode(ByVal strUnicode As String) As Byte()
Dim lLen As Long
Dim lBufferSize As Long
Dim lResult As Long
Dim byteArray() As Byte

lLen = Len(strUnicode)
If lLen = 0 Then
    ReDim UTF8_Encode(-1)
    Exit Function
End If

'Set buffer for longest possible string.
lBufferSize = lLen * 3 + 1
ReDim byteArray(lBufferSize - 1)
'Translate using code page 65001(UTF-8).
lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), lLen, byteArray(0), lBufferSize, vbNullString, 0)
'Trim result to actual length.
If lResult Then
    lResult = lResult - 1
    ReDim Preserve byteArray(lResult)
    UTF8_Encode = byteArray
Else
    ReDim UTF8_Encode(-1)
End If
End Function

'Purpose:Convert Unicode string to UTF-8.
Public Function UTF8Length(ByVal strUnicode As String) As Long
Dim lLen As Long
Dim lBufferSize As Long
Dim byteArray() As Byte

lLen = Len(strUnicode)
If lLen = 0 Then
    UTF8Length = 0
    Exit Function
End If

'Set buffer for longest possible string.
lBufferSize = lLen * 3 + 1
ReDim byteArray(lBufferSize - 1)
'Translate using code page 65001(UTF-8).
UTF8Length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), lLen, byteArray(0), lBufferSize, vbNullString, 0)
End Function

Public Function MaxUTF8Length(strUnicode As String, iCount As Integer) As String
Dim iBytePos As Integer
Dim iStringPos As Integer
Dim lResult As Long
Dim byteCharArray(4) As Byte

'Need separte pointers for string buffer position and byte array position

If Len(strUnicode) = 0 Then
    MaxUTF8Length = ""
    iCount = 0
    Exit Function
End If

iBytePos = 0
iStringPos = 1

Do
    'Translate using code page 65001(UTF-8).
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Mid$(strUnicode, iStringPos, 1)), 1, byteCharArray(0), 4, vbNullString, 0)
    'Trim result to actual length.
    If lResult > 0 Then
        If (iBytePos + lResult) > iCount Then
            iCount = iBytePos
            Exit Function
        Else
            MaxUTF8Length = MaxUTF8Length + Mid$(strUnicode, iStringPos, 1)
            iStringPos = iStringPos + 1
            iBytePos = iBytePos + lResult
        End If
    Else
        iCount = 0
        MaxUTF8Length = ""
        Exit Function
    End If
Loop
End Function
