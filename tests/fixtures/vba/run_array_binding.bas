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
Public Function ReadLong(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadVariant(ByRef values() As Variant) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadVariant = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadVariant = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadObject(ByRef values() As Object) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadObject = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadObject = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadBox(ByRef values() As Box) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadBox = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadBox = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadOther(ByRef values() As Other) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadOther = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadOther = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadKind(ByRef values() As Kind) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadKind = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadKind = "caught:" & CStr(Err.Number)
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
Public Function ReadLong(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadLong = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadLong = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadVariant(ByRef values() As Variant) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadVariant = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadVariant = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadObject(ByRef values() As Object) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadObject = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadObject = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadBox(ByRef values() As Box) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadBox = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadBox = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadOther(ByRef values() As Other) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadOther = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadOther = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReadKind(ByRef values() As Kind) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReadKind = TypeName(values) & ":" & CStr(VarType(values)) & ":" & CStr(LBound(values)) & ":" & TypeName(values(LBound(values)))
Exit Function
Caught:
ReadKind = "caught:" & CStr(Err.Number)
Err.Clear
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Enum Kind
FirstKind = 1
SecondKind = 2
End Enum
Public Function MakeLong() As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
MakeLong = a
End Function
Public Function MakeVariant() As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
MakeVariant = a
End Function
Public Function MakeObject() As Variant
Dim a() As Object
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
MakeObject = a
End Function
Public Function MakeBox() As Variant
Dim a() As Box
ReDim a(1 To 2)
Set a(1) = Nothing
Set a(2) = Nothing
MakeBox = a
End Function
Private Function Try0() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try1() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try2() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", (a))
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try3() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", MakeLong())
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try4() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try5() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a As Variant
a = CLng(1)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try6() As String
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
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try7() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try8() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try9() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try10() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadVariant", (a))
failure = Err.Number
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try11() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadVariant", MakeVariant())
failure = Err.Number
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try12() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try13() As String
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
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try14() As String
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
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try15() As String
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
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try16() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try17() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", (a))
failure = Err.Number
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try18() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadBox", MakeBox())
failure = Err.Number
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try19() As String
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
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try20() As String
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
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try21() As String
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
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try22() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Object
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try23() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadObject", MakeObject())
failure = Err.Number
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try24() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadKind", a)
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
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try26() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try27() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadKind", (a))
failure = Err.Number
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try28() As String
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
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try29() As String
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
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try30() As String
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
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try31() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try32() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try33() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try34() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", (a))
failure = Err.Number
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try35() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", MakeLong())
failure = Err.Number
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try36() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try37() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a As Variant
a = CLng(1)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try38() As String
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
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try39() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try40() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try41() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try42() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadVariant", (a))
failure = Err.Number
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try43() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadVariant", MakeVariant())
failure = Err.Number
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try44() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try45() As String
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
answer = app.Run("Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try47() As String
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
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try48() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
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
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", (a))
failure = Err.Number
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try50() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadBox", MakeBox())
failure = Err.Number
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
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
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try52() As String
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
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try53() As String
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
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try54() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Object
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try55() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadObject", MakeObject())
failure = Err.Number
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try56() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try57() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try58() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try59() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadKind", (a))
failure = Err.Number
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try60() As String
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
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try61() As String
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
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try62() As String
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
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try63() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReadOther", a)
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
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try64 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try65() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try65 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try66() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", (a))
failure = Err.Number
Err.Clear
Try66 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try67() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", MakeLong())
failure = Err.Number
Err.Clear
Try67 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try68() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try68 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try69() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a As Variant
a = CLng(1)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try69 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try70() As String
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
Try70 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try71() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadLong", a)
failure = Err.Number
Err.Clear
Try71 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try72() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try72 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try73() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try73 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try74() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadVariant", (a))
failure = Err.Number
Err.Clear
Try74 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try75() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadVariant", MakeVariant())
failure = Err.Number
Err.Clear
Try75 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try76() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadVariant", a)
failure = Err.Number
Err.Clear
Try76 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try77() As String
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
Try77 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try78() As String
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
Try78 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try79() As String
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
Try79 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try80() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", a)
failure = Err.Number
Err.Clear
Try80 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try81() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", (a))
failure = Err.Number
Err.Clear
Try81 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try82() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadBox", MakeBox())
failure = Err.Number
Err.Clear
Try82 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try83() As String
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
Try83 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try84() As String
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
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", a)
failure = Err.Number
Err.Clear
Try85 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try86() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Object
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
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
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadObject", MakeObject())
failure = Err.Number
Err.Clear
Try87 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try88() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try88 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try89() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try89 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try90() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadKind", a)
failure = Err.Number
Err.Clear
Try90 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try91() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadKind", (a))
failure = Err.Number
Err.Clear
Try91 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try92() As String
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
Try92 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try93() As String
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
Try93 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try94() As String
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
Try94 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try95() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReadOther", a)
failure = Err.Number
Err.Clear
Try95 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try96() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try96 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try97() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try97 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try98() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, (a))
failure = Err.Number
Err.Clear
Try98 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try99() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, MakeLong())
failure = Err.Number
Err.Clear
Try99 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try100() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try100 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try101() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a As Variant
a = CLng(1)
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
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
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try102 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try103() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadLong", VbMethod, a)
failure = Err.Number
Err.Clear
Try103 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try104() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadVariant", VbMethod, a)
failure = Err.Number
Err.Clear
Try104 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try105() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadVariant", VbMethod, a)
failure = Err.Number
Err.Clear
Try105 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try106() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Variant
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadVariant", VbMethod, (a))
failure = Err.Number
Err.Clear
Try106 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try107() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadVariant", VbMethod, MakeVariant())
failure = Err.Number
Err.Clear
Try107 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try108() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadVariant", VbMethod, a)
failure = Err.Number
Err.Clear
Try108 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try109() As String
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
Try109 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try110() As String
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
Try110 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try111() As String
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
Try111 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try112() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, a)
failure = Err.Number
Err.Clear
Try112 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try113() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Box
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, (a))
failure = Err.Number
Err.Clear
Try113 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try114() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadBox", VbMethod, MakeBox())
failure = Err.Number
Err.Clear
Try114 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try115() As String
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
Try115 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try116() As String
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
Try116 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try117() As String
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
Try117 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try118() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Object
ReDim original(1 To 2)
Set original(1) = Nothing
Set original(2) = Nothing
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, a)
failure = Err.Number
Err.Clear
Try118 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try119() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadObject", VbMethod, MakeObject())
failure = Err.Number
Err.Clear
Try119 = CStr(failure) & "|" & CStr(answer) & "|" & "rvalue" & "|" & Trace
End Function
Private Function Try120() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadKind", VbMethod, a)
failure = Err.Number
Err.Clear
Try120 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try121() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadKind", VbMethod, a)
failure = Err.Number
Err.Clear
Try121 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try122() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadKind", VbMethod, a)
failure = Err.Number
Err.Clear
Try122 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try123() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim original() As Kind
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(receiver, "ReadKind", VbMethod, (a))
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
Try125 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try126() As String
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
Try126 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try127() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Other
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
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadLong(a)
failure = Err.Number
Err.Clear
Try128 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try129() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Variant
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadVariant(a)
failure = Err.Number
Err.Clear
Try129 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try130() As String
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
answer = Targets.ReadBox(a)
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
answer = Targets.ReadObject(a)
failure = Err.Number
Err.Clear
Try131 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try132() As String
Dim app As Object
Dim receiver As ArrayReceiver
Dim answer As Variant
Dim failure As Long
Dim a() As Kind
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Set receiver = New ArrayReceiver
Trace = ""
On Error Resume Next
Err.Clear
answer = Targets.ReadKind(a)
failure = Err.Number
Err.Clear
Try132 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Private Function Try133() As String
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
answer = Targets.ReadOther(a)
failure = Err.Number
Err.Clear
Try133 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(VarType(a)) & ":" & TypeName(a) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "early_long_typed=" & Try0() & vbLf
report = report & "early_long_wrapped=" & Try1() & vbLf
report = report & "early_long_forced=" & Try2() & vbLf
report = report & "early_long_function=" & Try3() & vbLf
report = report & "early_long_wrong_variant=" & Try4() & vbLf
report = report & "early_long_scalar=" & Try5() & vbLf
report = report & "early_long_uninitialized=" & Try6() & vbLf
report = report & "early_long_from_kind=" & Try7() & vbLf
report = report & "early_variant_typed=" & Try8() & vbLf
report = report & "early_variant_wrapped=" & Try9() & vbLf
report = report & "early_variant_forced=" & Try10() & vbLf
report = report & "early_variant_function=" & Try11() & vbLf
report = report & "early_variant_wrong_long=" & Try12() & vbLf
report = report & "early_box_typed=" & Try13() & vbLf
report = report & "early_box_wrong_other=" & Try14() & vbLf
report = report & "early_box_from_object=" & Try15() & vbLf
report = report & "early_box_wrapped=" & Try16() & vbLf
report = report & "early_box_forced=" & Try17() & vbLf
report = report & "early_box_function=" & Try18() & vbLf
report = report & "early_object_typed=" & Try19() & vbLf
report = report & "early_object_from_box=" & Try20() & vbLf
report = report & "early_object_from_other=" & Try21() & vbLf
report = report & "early_object_wrapped=" & Try22() & vbLf
report = report & "early_object_function=" & Try23() & vbLf
report = report & "early_kind_typed=" & Try24() & vbLf
report = report & "early_kind_from_long=" & Try25() & vbLf
report = report & "early_kind_wrapped=" & Try26() & vbLf
report = report & "early_kind_forced=" & Try27() & vbLf
report = report & "early_other_typed=" & Try28() & vbLf
report = report & "early_other_wrong_box=" & Try29() & vbLf
report = report & "early_other_from_object=" & Try30() & vbLf
report = report & "early_other_uninitialized=" & Try31() & vbLf
report = report & "object_long_typed=" & Try32() & vbLf
report = report & "object_long_wrapped=" & Try33() & vbLf
report = report & "object_long_forced=" & Try34() & vbLf
report = report & "object_long_function=" & Try35() & vbLf
report = report & "object_long_wrong_variant=" & Try36() & vbLf
report = report & "object_long_scalar=" & Try37() & vbLf
report = report & "object_long_uninitialized=" & Try38() & vbLf
report = report & "object_long_from_kind=" & Try39() & vbLf
report = report & "object_variant_typed=" & Try40() & vbLf
report = report & "object_variant_wrapped=" & Try41() & vbLf
report = report & "object_variant_forced=" & Try42() & vbLf
report = report & "object_variant_function=" & Try43() & vbLf
report = report & "object_variant_wrong_long=" & Try44() & vbLf
report = report & "object_box_typed=" & Try45() & vbLf
report = report & "object_box_wrong_other=" & Try46() & vbLf
report = report & "object_box_from_object=" & Try47() & vbLf
report = report & "object_box_wrapped=" & Try48() & vbLf
report = report & "object_box_forced=" & Try49() & vbLf
report = report & "object_box_function=" & Try50() & vbLf
report = report & "object_object_typed=" & Try51() & vbLf
report = report & "object_object_from_box=" & Try52() & vbLf
report = report & "object_object_from_other=" & Try53() & vbLf
report = report & "object_object_wrapped=" & Try54() & vbLf
report = report & "object_object_function=" & Try55() & vbLf
report = report & "object_kind_typed=" & Try56() & vbLf
report = report & "object_kind_from_long=" & Try57() & vbLf
report = report & "object_kind_wrapped=" & Try58() & vbLf
report = report & "object_kind_forced=" & Try59() & vbLf
report = report & "object_other_typed=" & Try60() & vbLf
report = report & "object_other_wrong_box=" & Try61() & vbLf
report = report & "object_other_from_object=" & Try62() & vbLf
report = report & "object_other_uninitialized=" & Try63() & vbLf
report = report & "callbyname_long_typed=" & Try64() & vbLf
report = report & "callbyname_long_wrapped=" & Try65() & vbLf
report = report & "callbyname_long_forced=" & Try66() & vbLf
report = report & "callbyname_long_function=" & Try67() & vbLf
report = report & "callbyname_long_wrong_variant=" & Try68() & vbLf
report = report & "callbyname_long_scalar=" & Try69() & vbLf
report = report & "callbyname_long_uninitialized=" & Try70() & vbLf
report = report & "callbyname_long_from_kind=" & Try71() & vbLf
report = report & "callbyname_variant_typed=" & Try72() & vbLf
report = report & "callbyname_variant_wrapped=" & Try73() & vbLf
report = report & "callbyname_variant_forced=" & Try74() & vbLf
report = report & "callbyname_variant_function=" & Try75() & vbLf
report = report & "callbyname_variant_wrong_long=" & Try76() & vbLf
report = report & "callbyname_box_typed=" & Try77() & vbLf
report = report & "callbyname_box_wrong_other=" & Try78() & vbLf
report = report & "callbyname_box_from_object=" & Try79() & vbLf
report = report & "callbyname_box_wrapped=" & Try80() & vbLf
report = report & "callbyname_box_forced=" & Try81() & vbLf
report = report & "callbyname_box_function=" & Try82() & vbLf
report = report & "callbyname_object_typed=" & Try83() & vbLf
report = report & "callbyname_object_from_box=" & Try84() & vbLf
report = report & "callbyname_object_from_other=" & Try85() & vbLf
report = report & "callbyname_object_wrapped=" & Try86() & vbLf
report = report & "callbyname_object_function=" & Try87() & vbLf
report = report & "callbyname_kind_typed=" & Try88() & vbLf
report = report & "callbyname_kind_from_long=" & Try89() & vbLf
report = report & "callbyname_kind_wrapped=" & Try90() & vbLf
report = report & "callbyname_kind_forced=" & Try91() & vbLf
report = report & "callbyname_other_typed=" & Try92() & vbLf
report = report & "callbyname_other_wrong_box=" & Try93() & vbLf
report = report & "callbyname_other_from_object=" & Try94() & vbLf
report = report & "callbyname_other_uninitialized=" & Try95() & vbLf
report = report & "class_callbyname_long_typed=" & Try96() & vbLf
report = report & "class_callbyname_long_wrapped=" & Try97() & vbLf
report = report & "class_callbyname_long_forced=" & Try98() & vbLf
report = report & "class_callbyname_long_function=" & Try99() & vbLf
report = report & "class_callbyname_long_wrong_variant=" & Try100() & vbLf
report = report & "class_callbyname_long_scalar=" & Try101() & vbLf
report = report & "class_callbyname_long_uninitialized=" & Try102() & vbLf
report = report & "class_callbyname_long_from_kind=" & Try103() & vbLf
report = report & "class_callbyname_variant_typed=" & Try104() & vbLf
report = report & "class_callbyname_variant_wrapped=" & Try105() & vbLf
report = report & "class_callbyname_variant_forced=" & Try106() & vbLf
report = report & "class_callbyname_variant_function=" & Try107() & vbLf
report = report & "class_callbyname_variant_wrong_long=" & Try108() & vbLf
report = report & "class_callbyname_box_typed=" & Try109() & vbLf
report = report & "class_callbyname_box_wrong_other=" & Try110() & vbLf
report = report & "class_callbyname_box_from_object=" & Try111() & vbLf
report = report & "class_callbyname_box_wrapped=" & Try112() & vbLf
report = report & "class_callbyname_box_forced=" & Try113() & vbLf
report = report & "class_callbyname_box_function=" & Try114() & vbLf
report = report & "class_callbyname_object_typed=" & Try115() & vbLf
report = report & "class_callbyname_object_from_box=" & Try116() & vbLf
report = report & "class_callbyname_object_from_other=" & Try117() & vbLf
report = report & "class_callbyname_object_wrapped=" & Try118() & vbLf
report = report & "class_callbyname_object_function=" & Try119() & vbLf
report = report & "class_callbyname_kind_typed=" & Try120() & vbLf
report = report & "class_callbyname_kind_from_long=" & Try121() & vbLf
report = report & "class_callbyname_kind_wrapped=" & Try122() & vbLf
report = report & "class_callbyname_kind_forced=" & Try123() & vbLf
report = report & "class_callbyname_other_typed=" & Try124() & vbLf
report = report & "class_callbyname_other_wrong_box=" & Try125() & vbLf
report = report & "class_callbyname_other_from_object=" & Try126() & vbLf
report = report & "class_callbyname_other_uninitialized=" & Try127() & vbLf
report = report & "direct_long_typed=" & Try128() & vbLf
report = report & "direct_variant_typed=" & Try129() & vbLf
report = report & "direct_box_typed=" & Try130() & vbLf
report = report & "direct_object_typed=" & Try131() & vbLf
report = report & "direct_kind_typed=" & Try132() & vbLf
report = report & "direct_other_typed=" & Try133() & vbLf
Probe = report
End Function

