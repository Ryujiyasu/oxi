' SPDX-License-Identifier: CC-BY-SA-4.0
Option Explicit
Private hits As Long
Public Function FileAttributesIdentity() As Long
 FileAttributesIdentity = 2026102102
End Function
Public Function DirectoryEnvironment() As String
 DirectoryEnvironment = CurDir$
End Function
Private Function ParamByte(ByVal item As Byte) As Variant
 hits = hits + 1
 ParamByte = item
End Function
Private Function ParamInteger(ByVal item As Integer) As Variant
 hits = hits + 1
 ParamInteger = item
End Function
Private Function ParamLong(ByVal item As Long) As Variant
 hits = hits + 1
 ParamLong = item
End Function
Private Function ParamSingle(ByVal item As Single) As Variant
 hits = hits + 1
 ParamSingle = item
End Function
Private Function ParamDouble(ByVal item As Double) As Variant
 hits = hits + 1
 ParamDouble = item
End Function
Private Function ParamCurrency(ByVal item As Currency) As Variant
 hits = hits + 1
 ParamCurrency = item
End Function
Private Function ParamBoolean(ByVal item As Boolean) As Variant
 hits = hits + 1
 ParamBoolean = item
End Function
Private Function ParamString(ByVal item As String) As Variant
 hits = hits + 1
 ParamString = item
End Function
Private Function ParamDate(ByVal item As Date) As Variant
 hits = hits + 1
 ParamDate = item
End Function
Private Function ParamVariant(ByVal item As Variant) As Variant
 hits = hits + 1
 ParamVariant = item
End Function
Public Function ErrorProbe(ByVal kind As Long) As String
 Dim answer As Variant, description As String
 Dim localByte As Byte
 Dim localInteger As Integer
 Dim localLong As Long
 Dim localSingle As Single
 Dim localDouble As Double
 Dim localCurrency As Currency
 Dim localBoolean As Boolean
 Dim localString As String
 Dim localDate As Date
 hits = 0
 On Error GoTo Caught
 Select Case kind
 Case 0: localByte = CVErr(2042): answer = localByte
 Case 1: answer = CByte(CVErr(2042))
 Case 2: answer = ParamByte(CVErr(2042))
 Case 3: localInteger = CVErr(2042): answer = localInteger
 Case 4: answer = CInt(CVErr(2042))
 Case 5: answer = ParamInteger(CVErr(2042))
 Case 6: localLong = CVErr(2042): answer = localLong
 Case 7: answer = CLng(CVErr(2042))
 Case 8: answer = ParamLong(CVErr(2042))
 Case 9: localSingle = CVErr(2042): answer = localSingle
 Case 10: answer = CSng(CVErr(2042))
 Case 11: answer = ParamSingle(CVErr(2042))
 Case 12: localDouble = CVErr(2042): answer = localDouble
 Case 13: answer = CDbl(CVErr(2042))
 Case 14: answer = ParamDouble(CVErr(2042))
 Case 15: localCurrency = CVErr(2042): answer = localCurrency
 Case 16: answer = CCur(CVErr(2042))
 Case 17: answer = ParamCurrency(CVErr(2042))
 Case 18: localBoolean = CVErr(2042): answer = localBoolean
 Case 19: answer = CBool(CVErr(2042))
 Case 20: answer = ParamBoolean(CVErr(2042))
 Case 21: localString = CVErr(2042): answer = localString
 Case 22: answer = CStr(CVErr(2042))
 Case 23: answer = ParamString(CVErr(2042))
 Case 24: localDate = CVErr(2042): answer = localDate
 Case 25: answer = CDate(CVErr(2042))
 Case 26: answer = ParamDate(CVErr(2042))
 Case 27: answer = CVErr(2042)
 Case 28: answer = ParamVariant(CVErr(2042))
 End Select
 description = CStr(VarType(answer)) & "," & TypeName(answer)
 If VarType(answer) = vbDate Then answer = CDbl(answer)
 ErrorProbe = "value:" & description & "," & CStr(answer) & "|hits:" & CStr(hits)
 Exit Function
Caught:
 ErrorProbe = "error:" & CStr(Err.Number) & "|hits:" & CStr(hits)
End Function
