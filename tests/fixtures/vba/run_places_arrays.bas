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
Public Function PutText(ByRef value As String) As String
Driver.Trace = Driver.Trace & "T"
PutText = value
value = "abcdef"
PutText = PutText & ">" & value
End Function
Public Function PutLong(ByRef value As Long) As String
Driver.Trace = Driver.Trace & "T"
PutLong = CStr(value)
value = 99
PutLong = PutLong & ">" & CStr(value)
End Function
Public Function LongArray(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
values(LBound(values)) = 99
LongArray = CStr(LBound(values)) & ":" & CStr(UBound(values)) & ":" & CStr(values(LBound(values)))
Exit Function
Caught:
LongArray = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function VariantArray(ByRef values() As Variant) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
values(LBound(values)) = 99
VariantArray = CStr(LBound(values)) & ":" & CStr(UBound(values)) & ":" & CStr(values(LBound(values)))
Exit Function
Caught:
VariantArray = "caught:" & CStr(Err.Number)
Err.Clear
End Function
Public Function ReplaceLong(ByRef values() As Long) As String
On Error GoTo Caught
Driver.Trace = Driver.Trace & "T"
ReDim values(3 To 4)
values(3) = 8
ReplaceLong = CStr(LBound(values)) & ":" & CStr(UBound(values)) & ":" & CStr(values(LBound(values)))
Exit Function
Caught:
ReplaceLong = "caught:" & CStr(Err.Number)
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
Public Text As String
Public Count As Long
Public Data As Variant


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Type Datum
Text As String * 5
Count As Long
End Type
Private Function Try0() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutText", x)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try1() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String * 5
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutText", x)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try2() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutText", d.Text)
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Text) & "|" & Trace
End Function
Private Function Try3() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try4() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try5() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Long
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try6() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Integer
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try7() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Variant
x = CLng(1)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try8() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", d.Count)
failure = Err.Number
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Count) & "|" & Trace
End Function
Private Function Try9() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try10() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try11() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try12() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try13() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try14() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try15() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try16() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try17() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try18() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try19() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try20() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try21() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try22() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutText", x)
failure = Err.Number
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try23() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String * 5
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutText", x)
failure = Err.Number
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try24() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutText", d.Text)
failure = Err.Number
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Text) & "|" & Trace
End Function
Private Function Try25() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try26() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try27() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Long
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try28() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Integer
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try29() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Variant
x = CLng(1)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try30() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", d.Count)
failure = Err.Number
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Count) & "|" & Trace
End Function
Private Function Try31() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try32() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try33() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try34() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try35() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try36() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try37() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try38() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try39() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try40() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try41() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try42() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try43() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = app.Run("Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try44() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutText", x)
End With
failure = Err.Number
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try45() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String * 5
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutText", x)
End With
failure = Err.Number
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try46() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutText", d.Text)
End With
failure = Err.Number
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Text) & "|" & Trace
End Function
Private Function Try47() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutText", b.Text)
End With
failure = Err.Number
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try48() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutText", b.Text)
End With
failure = Err.Number
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try49() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Long
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", x)
End With
failure = Err.Number
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try50() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Integer
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", x)
End With
failure = Err.Number
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try51() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Variant
x = CLng(1)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", x)
End With
failure = Err.Number
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try52() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", d.Count)
End With
failure = Err.Number
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Count) & "|" & Trace
End Function
Private Function Try53() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", b.Count)
End With
failure = Err.Number
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try54() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", b.Count)
End With
failure = Err.Number
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try55() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", a(1))
End With
failure = Err.Number
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try56() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.PutLong", a(1))
End With
failure = Err.Number
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try57() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.LongArray", a)
End With
failure = Err.Number
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try58() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.LongArray", a)
End With
failure = Err.Number
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try59() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.LongArray", a)
End With
failure = Err.Number
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try60() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.VariantArray", a)
End With
failure = Err.Number
Err.Clear
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try61() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.VariantArray", a)
End With
failure = Err.Number
Err.Clear
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try62() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.ReplaceLong", a)
End With
failure = Err.Number
Err.Clear
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try63() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.ReplaceLong", a)
End With
failure = Err.Number
Err.Clear
Try63 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try64() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.VariantArray", b.Data)
End With
failure = Err.Number
Err.Clear
observed = b.Data
Try64 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try65() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
With app
answer = .Run("Targets.VariantArray", b.Data)
End With
failure = Err.Number
Err.Clear
observed = b.Data
Try65 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try66() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutText", x)
failure = Err.Number
Err.Clear
Try66 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try67() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As String * 5
x = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutText", x)
failure = Err.Number
Err.Clear
Try67 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try68() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutText", d.Text)
failure = Err.Number
Err.Clear
Try68 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Text) & "|" & Trace
End Function
Private Function Try69() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try69 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try70() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Text = "hi"
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutText", b.Text)
failure = Err.Number
Err.Clear
Try70 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Text) & "|" & Trace
End Function
Private Function Try71() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Long
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try71 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try72() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Integer
x = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try72 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try73() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim x As Variant
x = CLng(1)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", x)
failure = Err.Number
Err.Clear
Try73 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & "|" & Trace
End Function
Private Function Try74() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim d As Datum
d.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", d.Count)
failure = Err.Number
Err.Clear
Try74 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(d.Count) & "|" & Trace
End Function
Private Function Try75() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try75 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try76() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Count = 1
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", b.Count)
failure = Err.Number
Err.Clear
Try76 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(b.Count) & "|" & Trace
End Function
Private Function Try77() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try77 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try78() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.PutLong", a(1))
failure = Err.Number
Err.Clear
Try78 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(a(1)) & "|" & Trace
End Function
Private Function Try79() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try79 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try80() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try80 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try81() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.LongArray", a)
failure = Err.Number
Err.Clear
Try81 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try82() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Variant
ReDim a(1 To 2)
a(1) = CLng(1)
a(2) = CLng(2)
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try82 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try83() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.VariantArray", a)
failure = Err.Number
Err.Clear
Try83 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try84() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim a() As Long
ReDim a(1 To 2)
a(1) = 1
a(2) = 2
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try84 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try85() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim original() As Long
ReDim original(1 To 2)
original(1) = 1
original(2) = 2
Dim a As Variant
a = original
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ReplaceLong", a)
failure = Err.Number
Err.Clear
Try85 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(a(LBound(a))) & "|" & Trace
End Function
Private Function Try86() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Box
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try86 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Private Function Try87() As String
Dim app As Object
Dim answer As Variant
Dim failure As Long
Dim observed As Variant
Dim b As Object
Set b = New Box
b.Data = Array(CLng(1), CLng(2))
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.VariantArray", b.Data)
failure = Err.Number
Err.Clear
observed = b.Data
Try87 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(LBound(observed)) & ":" & CStr(UBound(observed)) & ":" & CStr(observed(LBound(observed))) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "early_string=" & Try0() & vbLf
report = report & "early_fixed_string=" & Try1() & vbLf
report = report & "early_record_string=" & Try2() & vbLf
report = report & "early_box_string=" & Try3() & vbLf
report = report & "early_late_box_string=" & Try4() & vbLf
report = report & "early_long=" & Try5() & vbLf
report = report & "early_integer=" & Try6() & vbLf
report = report & "early_variant_long=" & Try7() & vbLf
report = report & "early_record_long=" & Try8() & vbLf
report = report & "early_box_long=" & Try9() & vbLf
report = report & "early_late_box_long=" & Try10() & vbLf
report = report & "early_long_element=" & Try11() & vbLf
report = report & "early_variant_element=" & Try12() & vbLf
report = report & "early_long_array=" & Try13() & vbLf
report = report & "early_variant_long_array=" & Try14() & vbLf
report = report & "early_variant_array_wrong=" & Try15() & vbLf
report = report & "early_variant_array=" & Try16() & vbLf
report = report & "early_long_array_wrong=" & Try17() & vbLf
report = report & "early_replace_long_array=" & Try18() & vbLf
report = report & "early_replace_variant_long=" & Try19() & vbLf
report = report & "early_box_array=" & Try20() & vbLf
report = report & "early_late_box_array=" & Try21() & vbLf
report = report & "object_string=" & Try22() & vbLf
report = report & "object_fixed_string=" & Try23() & vbLf
report = report & "object_record_string=" & Try24() & vbLf
report = report & "object_box_string=" & Try25() & vbLf
report = report & "object_late_box_string=" & Try26() & vbLf
report = report & "object_long=" & Try27() & vbLf
report = report & "object_integer=" & Try28() & vbLf
report = report & "object_variant_long=" & Try29() & vbLf
report = report & "object_record_long=" & Try30() & vbLf
report = report & "object_box_long=" & Try31() & vbLf
report = report & "object_late_box_long=" & Try32() & vbLf
report = report & "object_long_element=" & Try33() & vbLf
report = report & "object_variant_element=" & Try34() & vbLf
report = report & "object_long_array=" & Try35() & vbLf
report = report & "object_variant_long_array=" & Try36() & vbLf
report = report & "object_variant_array_wrong=" & Try37() & vbLf
report = report & "object_variant_array=" & Try38() & vbLf
report = report & "object_long_array_wrong=" & Try39() & vbLf
report = report & "object_replace_long_array=" & Try40() & vbLf
report = report & "object_replace_variant_long=" & Try41() & vbLf
report = report & "object_box_array=" & Try42() & vbLf
report = report & "object_late_box_array=" & Try43() & vbLf
report = report & "with_object_string=" & Try44() & vbLf
report = report & "with_object_fixed_string=" & Try45() & vbLf
report = report & "with_object_record_string=" & Try46() & vbLf
report = report & "with_object_box_string=" & Try47() & vbLf
report = report & "with_object_late_box_string=" & Try48() & vbLf
report = report & "with_object_long=" & Try49() & vbLf
report = report & "with_object_integer=" & Try50() & vbLf
report = report & "with_object_variant_long=" & Try51() & vbLf
report = report & "with_object_record_long=" & Try52() & vbLf
report = report & "with_object_box_long=" & Try53() & vbLf
report = report & "with_object_late_box_long=" & Try54() & vbLf
report = report & "with_object_long_element=" & Try55() & vbLf
report = report & "with_object_variant_element=" & Try56() & vbLf
report = report & "with_object_long_array=" & Try57() & vbLf
report = report & "with_object_variant_long_array=" & Try58() & vbLf
report = report & "with_object_variant_array_wrong=" & Try59() & vbLf
report = report & "with_object_variant_array=" & Try60() & vbLf
report = report & "with_object_long_array_wrong=" & Try61() & vbLf
report = report & "with_object_replace_long_array=" & Try62() & vbLf
report = report & "with_object_replace_variant_long=" & Try63() & vbLf
report = report & "with_object_box_array=" & Try64() & vbLf
report = report & "with_object_late_box_array=" & Try65() & vbLf
report = report & "callbyname_string=" & Try66() & vbLf
report = report & "callbyname_fixed_string=" & Try67() & vbLf
report = report & "callbyname_record_string=" & Try68() & vbLf
report = report & "callbyname_box_string=" & Try69() & vbLf
report = report & "callbyname_late_box_string=" & Try70() & vbLf
report = report & "callbyname_long=" & Try71() & vbLf
report = report & "callbyname_integer=" & Try72() & vbLf
report = report & "callbyname_variant_long=" & Try73() & vbLf
report = report & "callbyname_record_long=" & Try74() & vbLf
report = report & "callbyname_box_long=" & Try75() & vbLf
report = report & "callbyname_late_box_long=" & Try76() & vbLf
report = report & "callbyname_long_element=" & Try77() & vbLf
report = report & "callbyname_variant_element=" & Try78() & vbLf
report = report & "callbyname_long_array=" & Try79() & vbLf
report = report & "callbyname_variant_long_array=" & Try80() & vbLf
report = report & "callbyname_variant_array_wrong=" & Try81() & vbLf
report = report & "callbyname_variant_array=" & Try82() & vbLf
report = report & "callbyname_long_array_wrong=" & Try83() & vbLf
report = report & "callbyname_replace_long_array=" & Try84() & vbLf
report = report & "callbyname_replace_variant_long=" & Try85() & vbLf
report = report & "callbyname_box_array=" & Try86() & vbLf
report = report & "callbyname_late_box_array=" & Try87() & vbLf
Probe = report
End Function

