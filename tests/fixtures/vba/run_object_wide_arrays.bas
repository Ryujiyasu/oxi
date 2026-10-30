VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Attribute VB_Name = "Targets"
Public Function ReadBox(ByRef values() As Box) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadBox = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadBox = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadOther(ByRef values() As Other) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadOther = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadOther = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadObject(ByRef values() As Object) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadObject = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadObject = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLong(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLongLong(ByRef values() As LongLong) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLongLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLongLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLongPtr(ByRef values() As LongPtr) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLongPtr = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLongPtr = "caught:" & CStr(Err.Number)
Err.Clear
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Tag As Long


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Tag As Long


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ReadBox(ByRef values() As Box) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadBox = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadBox = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadOther(ByRef values() As Other) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadOther = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadOther = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadObject(ByRef values() As Object) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadObject = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadObject = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLong(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLongLong(ByRef values() As LongLong) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLongLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLongLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadLongPtr(ByRef values() As LongPtr) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLongPtr = TypeName(values) & ":" & CStr(VarType(values)) & ":" & TypeName(values(LBound(values))) & ":" & TypeName(values(UBound(values)))
Exit Function
Caught:
ReadLongPtr = "caught:" & CStr(Err.Number)
Err.Clear
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Function MakeWide() As Variant
Dim a() As LongLong
ReDim a(1 To 2)
MakeWide = a
End Function
Private Function Try0() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try1() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try2() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try3() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try4() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try5() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try6() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try7() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try8() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try9() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try10() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try11() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try12() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try13() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try14() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try15() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try16() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try17() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try18() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try19() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try20() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try21() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try22() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try23() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try24() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try25() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try26() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try27() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try28() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try29() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try30() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try31() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try32() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try33() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try34() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try35() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try36() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try37() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", (a))
failure = Err.Number
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try38() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLongLong", MakeWide())
failure = Err.Number
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try39() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try40() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try41() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try42() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try43() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try44() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try45() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try46() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try47() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try48() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try49() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try50() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try51() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try52() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try53() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try54() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try55() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try56() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try57() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try58() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try59() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try60() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try61() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try62() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try63() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try63 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try64() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try64 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try65() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try65 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try66() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try66 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try67() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try67 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try68() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try68 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try69() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try69 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try70() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try70 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try71() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try71 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try72() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try72 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try73() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try73 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try74() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try74 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try75() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try75 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try76() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", (a))
failure = Err.Number
Err.Clear
Try76 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try77() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLongLong", MakeWide())
failure = Err.Number
Err.Clear
Try77 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try78() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try78 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try79() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try79 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try80() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try80 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try81() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try81 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try82() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try82 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try83() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try83 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try84() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try84 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try85() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try85 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try86() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try86 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try87() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try87 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try88() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try88 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try89() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try89 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try90() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try90 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try91() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try91 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try92() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try92 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try93() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try93 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try94() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try94 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try95() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try95 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try96() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try96 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try97() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try97 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try98() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try98 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try99() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try99 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try100() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try100 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try101() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try101 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try102() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try102 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try103() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try103 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try104() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try104 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try105() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try105 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try106() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try106 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try107() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try107 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try108() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try108 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try109() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try109 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try110() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try110 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try111() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try111 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try112() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try112 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try113() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongPtr", a)
failure = Err.Number
Err.Clear
Try113 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try114() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", a)
failure = Err.Number
Err.Clear
Try114 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try115() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", (a))
failure = Err.Number
Err.Clear
Try115 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try116() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLongLong", MakeWide())
failure = Err.Number
Err.Clear
Try116 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try117() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try117 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try118() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try118 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try119() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try119 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try120() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try120 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try121() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try121 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try122() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try122 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try123() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try123 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try124() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try124 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try125() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try125 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try126() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try126 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try127() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try127 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try128() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try128 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try129() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try129 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try130() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try130 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try131() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try131 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try132() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try132 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try133() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try133 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try134() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try134 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try135() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try135 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try136() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try136 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try137() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try137 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try138() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try138 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try139() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadOther", VbMethod, a)
failure = Err.Number
Err.Clear
Try139 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try140() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try140 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try141() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try141 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try142() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try142 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try143() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongPtr", VbMethod, a)
failure = Err.Number
Err.Clear
Try143 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try144() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try144 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try145() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try145 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try146() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try146 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try147() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongPtr", VbMethod, a)
failure = Err.Number
Err.Clear
Try147 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try148() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongLong
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try148 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try149() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try149 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try150() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try150 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try151() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
ReDim a(1 To 2)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongPtr", VbMethod, a)
failure = Err.Number
Err.Clear
Try151 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try152() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As LongPtr
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongPtr", VbMethod, a)
failure = Err.Number
Err.Clear
Try152 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try153() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try153 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try154() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As LongLong
ReDim original(1 To 2)
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, (a))
failure = Err.Number
Err.Clear
Try154 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try155() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLongLong", VbMethod, MakeWide())
failure = Err.Number
Err.Clear
Try155 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try156() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Box
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadBox(a)
failure = Err.Number
Err.Clear
Try156 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try157() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
ReDim a(1 To 2)
Set a(1) = New Other
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadOther(a)
failure = Err.Number
Err.Clear
Try157 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try158() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = New Box
Set a(2) = New Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadObject(a)
failure = Err.Number
Err.Clear
Try158 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "early_Box_from_Box_nothing=" & Try0() & vbLf
report = report & "early_Other_from_Box_nothing=" & Try1() & vbLf
report = report & "early_Object_from_Box_nothing=" & Try2() & vbLf
report = report & "early_Box_from_Box_box=" & Try3() & vbLf
report = report & "early_Other_from_Box_box=" & Try4() & vbLf
report = report & "early_Object_from_Box_box=" & Try5() & vbLf
report = report & "early_Box_from_Other_nothing=" & Try6() & vbLf
report = report & "early_Other_from_Other_nothing=" & Try7() & vbLf
report = report & "early_Object_from_Other_nothing=" & Try8() & vbLf
report = report & "early_Box_from_Other_other=" & Try9() & vbLf
report = report & "early_Other_from_Other_other=" & Try10() & vbLf
report = report & "early_Object_from_Other_other=" & Try11() & vbLf
report = report & "early_Box_from_Object_nothing=" & Try12() & vbLf
report = report & "early_Other_from_Object_nothing=" & Try13() & vbLf
report = report & "early_Object_from_Object_nothing=" & Try14() & vbLf
report = report & "early_Box_from_Object_box=" & Try15() & vbLf
report = report & "early_Other_from_Object_box=" & Try16() & vbLf
report = report & "early_Object_from_Object_box=" & Try17() & vbLf
report = report & "early_Box_from_Object_other=" & Try18() & vbLf
report = report & "early_Other_from_Object_other=" & Try19() & vbLf
report = report & "early_Object_from_Object_other=" & Try20() & vbLf
report = report & "early_Box_from_Object_mixed=" & Try21() & vbLf
report = report & "early_Other_from_Object_mixed=" & Try22() & vbLf
report = report & "early_Object_from_Object_mixed=" & Try23() & vbLf
report = report & "early_Long_from_Long_zeros=" & Try24() & vbLf
report = report & "early_LongLong_from_Long_zeros=" & Try25() & vbLf
report = report & "early_LongPtr_from_Long_zeros=" & Try26() & vbLf
report = report & "early_Long_uninitialized=" & Try27() & vbLf
report = report & "early_Long_from_LongLong_zeros=" & Try28() & vbLf
report = report & "early_LongLong_from_LongLong_zeros=" & Try29() & vbLf
report = report & "early_LongPtr_from_LongLong_zeros=" & Try30() & vbLf
report = report & "early_LongLong_uninitialized=" & Try31() & vbLf
report = report & "early_Long_from_LongPtr_zeros=" & Try32() & vbLf
report = report & "early_LongLong_from_LongPtr_zeros=" & Try33() & vbLf
report = report & "early_LongPtr_from_LongPtr_zeros=" & Try34() & vbLf
report = report & "early_LongPtr_uninitialized=" & Try35() & vbLf
report = report & "early_LongLong_wrapped=" & Try36() & vbLf
report = report & "early_LongLong_forced=" & Try37() & vbLf
report = report & "early_LongLong_function=" & Try38() & vbLf
report = report & "object_Box_from_Box_nothing=" & Try39() & vbLf
report = report & "object_Other_from_Box_nothing=" & Try40() & vbLf
report = report & "object_Object_from_Box_nothing=" & Try41() & vbLf
report = report & "object_Box_from_Box_box=" & Try42() & vbLf
report = report & "object_Other_from_Box_box=" & Try43() & vbLf
report = report & "object_Object_from_Box_box=" & Try44() & vbLf
report = report & "object_Box_from_Other_nothing=" & Try45() & vbLf
report = report & "object_Other_from_Other_nothing=" & Try46() & vbLf
report = report & "object_Object_from_Other_nothing=" & Try47() & vbLf
report = report & "object_Box_from_Other_other=" & Try48() & vbLf
report = report & "object_Other_from_Other_other=" & Try49() & vbLf
report = report & "object_Object_from_Other_other=" & Try50() & vbLf
report = report & "object_Box_from_Object_nothing=" & Try51() & vbLf
report = report & "object_Other_from_Object_nothing=" & Try52() & vbLf
report = report & "object_Object_from_Object_nothing=" & Try53() & vbLf
report = report & "object_Box_from_Object_box=" & Try54() & vbLf
report = report & "object_Other_from_Object_box=" & Try55() & vbLf
report = report & "object_Object_from_Object_box=" & Try56() & vbLf
report = report & "object_Box_from_Object_other=" & Try57() & vbLf
report = report & "object_Other_from_Object_other=" & Try58() & vbLf
report = report & "object_Object_from_Object_other=" & Try59() & vbLf
report = report & "object_Box_from_Object_mixed=" & Try60() & vbLf
report = report & "object_Other_from_Object_mixed=" & Try61() & vbLf
report = report & "object_Object_from_Object_mixed=" & Try62() & vbLf
report = report & "object_Long_from_Long_zeros=" & Try63() & vbLf
report = report & "object_LongLong_from_Long_zeros=" & Try64() & vbLf
report = report & "object_LongPtr_from_Long_zeros=" & Try65() & vbLf
report = report & "object_Long_uninitialized=" & Try66() & vbLf
report = report & "object_Long_from_LongLong_zeros=" & Try67() & vbLf
report = report & "object_LongLong_from_LongLong_zeros=" & Try68() & vbLf
report = report & "object_LongPtr_from_LongLong_zeros=" & Try69() & vbLf
report = report & "object_LongLong_uninitialized=" & Try70() & vbLf
report = report & "object_Long_from_LongPtr_zeros=" & Try71() & vbLf
report = report & "object_LongLong_from_LongPtr_zeros=" & Try72() & vbLf
report = report & "object_LongPtr_from_LongPtr_zeros=" & Try73() & vbLf
report = report & "object_LongPtr_uninitialized=" & Try74() & vbLf
report = report & "object_LongLong_wrapped=" & Try75() & vbLf
report = report & "object_LongLong_forced=" & Try76() & vbLf
report = report & "object_LongLong_function=" & Try77() & vbLf
report = report & "callbyname_Box_from_Box_nothing=" & Try78() & vbLf
report = report & "callbyname_Other_from_Box_nothing=" & Try79() & vbLf
report = report & "callbyname_Object_from_Box_nothing=" & Try80() & vbLf
report = report & "callbyname_Box_from_Box_box=" & Try81() & vbLf
report = report & "callbyname_Other_from_Box_box=" & Try82() & vbLf
report = report & "callbyname_Object_from_Box_box=" & Try83() & vbLf
report = report & "callbyname_Box_from_Other_nothing=" & Try84() & vbLf
report = report & "callbyname_Other_from_Other_nothing=" & Try85() & vbLf
report = report & "callbyname_Object_from_Other_nothing=" & Try86() & vbLf
report = report & "callbyname_Box_from_Other_other=" & Try87() & vbLf
report = report & "callbyname_Other_from_Other_other=" & Try88() & vbLf
report = report & "callbyname_Object_from_Other_other=" & Try89() & vbLf
report = report & "callbyname_Box_from_Object_nothing=" & Try90() & vbLf
report = report & "callbyname_Other_from_Object_nothing=" & Try91() & vbLf
report = report & "callbyname_Object_from_Object_nothing=" & Try92() & vbLf
report = report & "callbyname_Box_from_Object_box=" & Try93() & vbLf
report = report & "callbyname_Other_from_Object_box=" & Try94() & vbLf
report = report & "callbyname_Object_from_Object_box=" & Try95() & vbLf
report = report & "callbyname_Box_from_Object_other=" & Try96() & vbLf
report = report & "callbyname_Other_from_Object_other=" & Try97() & vbLf
report = report & "callbyname_Object_from_Object_other=" & Try98() & vbLf
report = report & "callbyname_Box_from_Object_mixed=" & Try99() & vbLf
report = report & "callbyname_Other_from_Object_mixed=" & Try100() & vbLf
report = report & "callbyname_Object_from_Object_mixed=" & Try101() & vbLf
report = report & "callbyname_Long_from_Long_zeros=" & Try102() & vbLf
report = report & "callbyname_LongLong_from_Long_zeros=" & Try103() & vbLf
report = report & "callbyname_LongPtr_from_Long_zeros=" & Try104() & vbLf
report = report & "callbyname_Long_uninitialized=" & Try105() & vbLf
report = report & "callbyname_Long_from_LongLong_zeros=" & Try106() & vbLf
report = report & "callbyname_LongLong_from_LongLong_zeros=" & Try107() & vbLf
report = report & "callbyname_LongPtr_from_LongLong_zeros=" & Try108() & vbLf
report = report & "callbyname_LongLong_uninitialized=" & Try109() & vbLf
report = report & "callbyname_Long_from_LongPtr_zeros=" & Try110() & vbLf
report = report & "callbyname_LongLong_from_LongPtr_zeros=" & Try111() & vbLf
report = report & "callbyname_LongPtr_from_LongPtr_zeros=" & Try112() & vbLf
report = report & "callbyname_LongPtr_uninitialized=" & Try113() & vbLf
report = report & "callbyname_LongLong_wrapped=" & Try114() & vbLf
report = report & "callbyname_LongLong_forced=" & Try115() & vbLf
report = report & "callbyname_LongLong_function=" & Try116() & vbLf
report = report & "class_callbyname_Box_from_Box_nothing=" & Try117() & vbLf
report = report & "class_callbyname_Other_from_Box_nothing=" & Try118() & vbLf
report = report & "class_callbyname_Object_from_Box_nothing=" & Try119() & vbLf
report = report & "class_callbyname_Box_from_Box_box=" & Try120() & vbLf
report = report & "class_callbyname_Other_from_Box_box=" & Try121() & vbLf
report = report & "class_callbyname_Object_from_Box_box=" & Try122() & vbLf
report = report & "class_callbyname_Box_from_Other_nothing=" & Try123() & vbLf
report = report & "class_callbyname_Other_from_Other_nothing=" & Try124() & vbLf
report = report & "class_callbyname_Object_from_Other_nothing=" & Try125() & vbLf
report = report & "class_callbyname_Box_from_Other_other=" & Try126() & vbLf
report = report & "class_callbyname_Other_from_Other_other=" & Try127() & vbLf
report = report & "class_callbyname_Object_from_Other_other=" & Try128() & vbLf
report = report & "class_callbyname_Box_from_Object_nothing=" & Try129() & vbLf
report = report & "class_callbyname_Other_from_Object_nothing=" & Try130() & vbLf
report = report & "class_callbyname_Object_from_Object_nothing=" & Try131() & vbLf
report = report & "class_callbyname_Box_from_Object_box=" & Try132() & vbLf
report = report & "class_callbyname_Other_from_Object_box=" & Try133() & vbLf
report = report & "class_callbyname_Object_from_Object_box=" & Try134() & vbLf
report = report & "class_callbyname_Box_from_Object_other=" & Try135() & vbLf
report = report & "class_callbyname_Other_from_Object_other=" & Try136() & vbLf
report = report & "class_callbyname_Object_from_Object_other=" & Try137() & vbLf
report = report & "class_callbyname_Box_from_Object_mixed=" & Try138() & vbLf
report = report & "class_callbyname_Other_from_Object_mixed=" & Try139() & vbLf
report = report & "class_callbyname_Object_from_Object_mixed=" & Try140() & vbLf
report = report & "class_callbyname_Long_from_Long_zeros=" & Try141() & vbLf
report = report & "class_callbyname_LongLong_from_Long_zeros=" & Try142() & vbLf
report = report & "class_callbyname_LongPtr_from_Long_zeros=" & Try143() & vbLf
report = report & "class_callbyname_Long_uninitialized=" & Try144() & vbLf
report = report & "class_callbyname_Long_from_LongLong_zeros=" & Try145() & vbLf
report = report & "class_callbyname_LongLong_from_LongLong_zeros=" & Try146() & vbLf
report = report & "class_callbyname_LongPtr_from_LongLong_zeros=" & Try147() & vbLf
report = report & "class_callbyname_LongLong_uninitialized=" & Try148() & vbLf
report = report & "class_callbyname_Long_from_LongPtr_zeros=" & Try149() & vbLf
report = report & "class_callbyname_LongLong_from_LongPtr_zeros=" & Try150() & vbLf
report = report & "class_callbyname_LongPtr_from_LongPtr_zeros=" & Try151() & vbLf
report = report & "class_callbyname_LongPtr_uninitialized=" & Try152() & vbLf
report = report & "class_callbyname_LongLong_wrapped=" & Try153() & vbLf
report = report & "class_callbyname_LongLong_forced=" & Try154() & vbLf
report = report & "class_callbyname_LongLong_function=" & Try155() & vbLf
report = report & "direct_Box_from_Box_box=" & Try156() & vbLf
report = report & "direct_Other_from_Other_other=" & Try157() & vbLf
report = report & "direct_Object_from_Object_mixed=" & Try158() & vbLf
Probe = report
End Function

