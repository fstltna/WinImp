Option Strict Off
Option Explicit On
Module UTF8
	
	'Change Log:
	'220304 rjk: Created to support UTF8 communication with the server
	'301204 rjk: Added to remove SO/SI characters from the server used to indicate
	'            highlighting in UTF8 mode.
	'260505 rjk: Switched UTF8 from an registry option to login option.
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Integer, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer) As Integer
	
	Public Const CP_UTF8 As Integer = 65001
	
	'Purpose:Convert Utf8 to Unicode
	Public Function UTF8_Decode(ByRef baUTF8() As Byte) As String
		Dim lUtf8Size As Integer
		Dim strBuffer As String
		Dim lBufferSize As Integer
		Dim lResult As Integer
		Dim lPos As Integer
		
		UTF8_Decode = vbNullString
		
		lUtf8Size = UBound(baUTF8) + 1
		If lUtf8Size <= 0 Then Exit Function
		
		'Set buffer for longest possible string i.e. each byte is
		'ANSI, thus 1 unicode(2 bytes)for every utf-8 character.
		lBufferSize = lUtf8Size * 2
		strBuffer = New String(vbNullChar, lBufferSize)
		'Translate using code page 65001(UTF-8)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lResult = MultiByteToWideChar(CP_UTF8, 0, baUTF8(0), lUtf8Size, StrPtr(strBuffer), lBufferSize)
		strBuffer = Left(strBuffer, lResult)
		'301204 rjk: Added to remove SO/SI characters from the server used to indicate
		'            highlighting in UTF8 mode.
		lPos = 1
		While InStr(lPos, strBuffer, Chr(14)) > 0
			lPos = InStr(strBuffer, Chr(14))
			strBuffer = Left(strBuffer, lPos - 1) & Right(strBuffer, Len(strBuffer) - lPos)
		End While
		lPos = 1
		While InStr(lPos, strBuffer, Chr(15)) > 0
			lPos = InStr(strBuffer, Chr(15))
			strBuffer = Left(strBuffer, lPos - 1) & Right(strBuffer, Len(strBuffer) - lPos)
		End While
		'Trim result to actual length
		If lResult Then
			UTF8_Decode = strBuffer
		End If
	End Function
	
	'Purpose:Convert Unicode string to UTF-8.
	Public Function UTF8_Encode(ByVal strUnicode As String) As Byte()
		Dim lLen As Integer
		Dim lBufferSize As Integer
		Dim lResult As Integer
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
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), lLen, byteArray(0), lBufferSize, vbNullString, 0)
		'Trim result to actual length.
		If lResult Then
			lResult = lResult - 1
			ReDim Preserve byteArray(lResult)
			UTF8_Encode = VB6.CopyArray(byteArray)
		Else
			ReDim UTF8_Encode(-1)
		End If
	End Function
	
	'Purpose:Convert Unicode string to UTF-8.
	Public Function UTF8Length(ByVal strUnicode As String) As Integer
		Dim lLen As Integer
		Dim lBufferSize As Integer
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
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		UTF8Length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), lLen, byteArray(0), lBufferSize, vbNullString, 0)
	End Function
	
	Public Function MaxUTF8Length(ByRef strUnicode As String, ByRef iCount As Short) As String
		Dim iBytePos As Short
		Dim iStringPos As Short
		Dim lResult As Integer
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
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Mid(strUnicode, iStringPos, 1)), 1, byteCharArray(0), 4, vbNullString, 0)
			'Trim result to actual length.
			If lResult > 0 Then
				If (iBytePos + lResult) > iCount Then
					iCount = iBytePos
					Exit Function
				Else
					MaxUTF8Length = MaxUTF8Length & Mid(strUnicode, iStringPos, 1)
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
End Module