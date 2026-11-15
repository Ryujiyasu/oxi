' SPDX-License-Identifier: CC-BY-SA-4.0
Attribute VB_Name = "VarPtrIdentity"
Option Explicit

Private Type PointerRecord
 A As Integer
 B As Long
 C As LongPtr
End Type
Private GlobalLong As Long
Private Function PV(ByRef x As Variant) As LongPtr
 PV = VarPtr(x)
End Function
Private Function PVVal(ByVal x As Variant) As LongPtr
 PVVal = VarPtr(x)
End Function
Private Function PL(ByRef x As Long) As LongPtr
 PL = VarPtr(x)
End Function
Private Function PLVal(ByVal x As Long) As LongPtr
 PLVal = VarPtr(x)
End Function
Private Function PI(ByRef x As Integer) As LongPtr
 PI = VarPtr(x)
End Function
Private Function PS(ByRef x As String) As LongPtr
 PS = VarPtr(x)
End Function
Private Function PSVal(ByVal x As String) As LongPtr
 PSVal = VarPtr(x)
End Function
Private Function PG() As LongPtr
 PG = VarPtr(GlobalLong)
End Function
Private Function PST() As LongPtr
 Static x As Long
 PST = VarPtr(x)
End Function
Public Function Probe(ByVal caseID As Long, ByVal binding As Long) As Variant
 Dim v As Variant, w As Variant, result As Variant
 Dim l As Long, p As LongPtr, q As LongPtr, s As String, wasNull As Boolean
 Dim ai(0 To 2) As Integer, al(0 To 2) As Long, av(0 To 2) As Variant
 Dim gridValues(-1 To 1, 3 To 4) As Integer
 Dim objectValue As Object, fixedText As String * 3
 Dim rr(0 To 1) As PointerRecord
 On Error Resume Next
 Err.Clear
 Select Case caseID
 Case 0
  p = VarPtr(v): q = VarPtr(w): v = 42
  result = Array(0, p <> 0, p = VarPtr(v), p <> q, PV(v) = p, PVVal(v) <> p, p = VarPtr(v), VarType(VarPtr(v)))
 Case 1
  l = 42: p = VarPtr(l)
  result = Array(0, p <> 0, PL(l) = p, PLVal(l) <> p, PV(l) = p, PVVal(l) <> p, VarType(VarPtr(l)), TypeName(VarPtr(l)))
 Case 2
  p = VarPtr(s): wasNull = (StrPtr(s) = 0): s = "a"
  result = Array(0, p <> 0, wasNull, StrPtr(s) <> 0, p = VarPtr(s), p <> StrPtr(s), PS(s) = p, PSVal(s) <> p)
 Case 3
  p = VarPtr(ai(0))
  result = Array(0, p <> 0, p = VarPtr(ai(0)), VarPtr(ai(1)) - p, VarPtr(ai(2)) - p, PI(ai(0)) = p, VarType(VarPtr(ai(0))), TypeName(VarPtr(ai(0))))
 Case 4
  p = VarPtr(al(0))
  result = Array(0, p <> 0, p = VarPtr(al(0)), VarPtr(al(1)) - p, VarPtr(al(2)) - p, PL(al(0)) = p, VarType(VarPtr(al(0))), TypeName(VarPtr(al(0))))
 Case 5
  p = VarPtr(av(0))
  result = Array(0, p <> 0, p = VarPtr(av(0)), VarPtr(av(1)) - p, VarPtr(av(2)) - p, PV(av(0)) = p, VarType(VarPtr(av(0))), TypeName(VarPtr(av(0))))
 Case 6
  p = VarPtr(rr(0))
  result = Array(0, p <> 0, VarPtr(rr(0).B) - p, VarPtr(rr(0).C) - p, VarPtr(rr(1)) - p, PL(rr(0).B) = VarPtr(rr(0).B), VarType(VarPtr(rr(0))), TypeName(VarPtr(rr(0))))
 Case 7
  result = Array(0, PG() = PG(), PST() = PST(), PG() <> PST(), VarType(PG()), TypeName(PST()), Empty, Empty)
 Case 8
  v = 1: p = VarPtr(v)
  result = Array(0, p = VBA.VarPtr(v), p = varptr(v), VarType(VBA.VarPtr(v)), TypeName(VBA.VarPtr(v)), Empty, Empty, Empty)
 Case 9
  l = 42: p = VarPtr(l)
  result = Array(0, PL((l)) <> p, VarPtr((l)) = p, VarType(VarPtr((l))), TypeName(VarPtr((l))), Empty, Empty, Empty)
 Case 10
  p = VarPtr(gridValues(-1, 3))
  result = Array(0, VarPtr(gridValues(0, 3)) - p, VarPtr(gridValues(-1, 4)) - p, VarPtr(gridValues(1, 4)) - p, PI(gridValues(0, 3)) = VarPtr(gridValues(0, 3)), VarType(VarPtr(gridValues(0, 3))), TypeName(VarPtr(gridValues(0, 3))), Empty)
 Case 11
  Set objectValue = New Collection: p = VarPtr(objectValue)
  result = Array(0, p <> 0, p = VarPtr(objectValue), p <> ObjPtr(objectValue), PO(objectValue) = p, VarType(VarPtr(objectValue)), TypeName(VarPtr(objectValue)), Empty)
 Case 12
  p = VarPtr(fixedText)
  result = Array(0, p <> 0, p = StrPtr(fixedText), p = VarPtr(fixedText), PS(fixedText) = p, VarType(VarPtr(fixedText)), TypeName(VarPtr(fixedText)), Empty)
 Case 13
  p = VarPtr(42): q = VarPtr(42)
  result = Array(0, p <> 0, p = q, VarPtr(42) = VarPtr(43), VarPtr(l + 1) = VarPtr(l + 2), VarPtr((l)) = VarPtr((l)), VarType(VarPtr(42)), TypeName(VarPtr(42)))
 Case 14
  GlobalLong = 42: result = PVWrite(GlobalLong)
 Case 15
  Const literalLong As Long = 42
  p = VarPtr(literalLong)
  result = Array(0, p = VarPtr(literalLong), p = VarPtr(42&), VarPtr(42) = VarPtr(42&), VarPtr(42#) = VarPtr(42#), VarPtr("a") = VarPtr("a"), VarPtr(-42) = VarPtr(-42), VarPtr(True) = VarPtr(True))
 Case 16
  GlobalLong = 42: result = PVWriteError(GlobalLong)
 Case 17
  Const constDouble As Double = 42
  Const constText As String = "a"
  Const constBoolean As Boolean = True
  result = Array(0, VarPtr(constDouble) = VarPtr(constDouble), VarPtr(constText) = VarPtr(constText), VarPtr(constBoolean) = VarPtr(constBoolean), VarPtr(constDouble) = VarPtr(42#), VarPtr(constText) = VarPtr("a"), VarPtr(constBoolean) = VarPtr(True), Empty)
 Case 18
  p = VarPtr(42): q = VarPtr(43)
  result = Array(0, p = q, p = VarPtr(42), p = VarPtr(43), VarPtr(42) = VarPtr(42), VarPtr(42) = VarPtr(43), VarType(p), Empty)
 Case 19
  p = VarPtr(42&): q = VarPtr(43&)
  result = Array(0, p = q, p = VarPtr(42&), p = VarPtr(43&), VarPtr(42&) = VarPtr(42&), VarPtr(42) = VarPtr(42&), VarType(p), Empty)
 Case 20
  p = VarPtr(42#): q = VarPtr(43#)
  result = Array(0, p = q, p = VarPtr(42#), p = VarPtr(43#), VarPtr(42#) = VarPtr(42#), VarPtr(42!) = VarPtr(42!), VarType(p), Empty)
 End Select
 result(0) = Err.Number
 Probe = result
End Function


Private Function PO(ByRef item As Object) As LongPtr
 PO = VarPtr(item)
End Function

Private Function PVWrite(ByRef x As Variant) As Variant
 Dim ptr As LongPtr
 ptr = VarPtr(x): x = 43
 PVWrite = Array(0, ptr <> VarPtr(GlobalLong), ptr = PV(x), GlobalLong, x, VarType(x), VarType(GlobalLong), ptr = VarPtr(x))
End Function

Private Function PVWriteError(ByRef x As Variant) As Variant
 Dim ptr As LongPtr
 On Error Resume Next
 ptr = VarPtr(x): x = "q"
 PVWriteError = Array(0, Err.Number, GlobalLong, x, VarType(x), VarType(GlobalLong), ptr = VarPtr(x), ptr = PV(x))
End Function
