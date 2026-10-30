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
Public Function Pair(ByRef first As Variant, ByRef second As Variant) As String
Driver.Trace = Driver.Trace & "T"
first = first + 10
second = first + 1
Pair = CStr(first) & ":" & CStr(second)
End Function
Public Function Opt(Optional ByRef first As Variant, Optional ByRef second As Variant) As String
Driver.Trace = Driver.Trace & "T"
Opt = CStr(IsMissing(first)) & ":" & CStr(IsMissing(second))
If Not IsMissing(first) Then first = 10
If Not IsMissing(second) Then second = 20
End Function
Public Function Whole(ByRef value As Long) As String
Driver.Trace = Driver.Trace & "T"
value = 99
Whole = "body"
End Function
Public Function Tail(ParamArray values() As Variant) As String
Driver.Trace = Driver.Trace & "T"
values(UBound(values)) = 88
Tail = CStr(UBound(values)) & ":" & CStr(values(UBound(values)))
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Function LateApp() As Object
Set LateApp = Application
End Function
Private Function Try0() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=y, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try1() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try2() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=(x), Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try3() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=(x))
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try4() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Opt")
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try5() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Opt", Arg1:=y)
failure = Err.Number
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try6() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg1:=n, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try7() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg1:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try8() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg1:=m, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try9() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg1:=CLng(1), Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try10() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=y, Macro:="Targets.Whole", Arg1:=x)
failure = Err.Number
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try11() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try12() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg2:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try13() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(Arg30:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try14() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=y, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try15() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try16() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=(x), Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try17() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=(x))
failure = Err.Number
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try18() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Opt")
failure = Err.Number
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try19() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Opt", Arg1:=y)
failure = Err.Number
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try20() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg1:=n, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try21() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg1:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try22() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg1:=m, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try23() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg1:=CLng(1), Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try24() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=y, Macro:="Targets.Whole", Arg1:=x)
failure = Err.Number
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try25() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try26() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg2:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try27() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg30:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try28() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(Arg01:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try29() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=y, Macro:="Targets.Pair", Arg1:=x)
End With
failure = Err.Number
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try30() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=x)
End With
failure = Err.Number
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try31() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=(x), Macro:="Targets.Pair", Arg1:=x)
End With
failure = Err.Number
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try32() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=(x))
End With
failure = Err.Number
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try33() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Opt")
End With
failure = Err.Number
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try34() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Opt", Arg1:=y)
End With
failure = Err.Number
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try35() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg1:=n, Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try36() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg1:=x, Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try37() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg1:=m, Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try38() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg1:=CLng(1), Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try39() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=y, Macro:="Targets.Whole", Arg1:=x)
End With
failure = Err.Number
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try40() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try41() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg2:=x, Macro:="Targets.Tail")
End With
failure = Err.Number
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try42() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg30:=x, Macro:="Targets.Tail")
End With
failure = Err.Number
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try43() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(Arg01:=x, Macro:="Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try44() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=y, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try45() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try46() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=(x), Macro:="Targets.Pair", Arg1:=x)
failure = Err.Number
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try47() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Pair", Arg1:=(x))
failure = Err.Number
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try48() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Opt")
failure = Err.Number
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try49() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Opt", Arg1:=y)
failure = Err.Number
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try50() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg1:=n, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try51() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg1:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try52() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg1:=m, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try53() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg1:=CLng(1), Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try54() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=y, Macro:="Targets.Whole", Arg1:=x)
failure = Err.Number
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try55() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try56() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg2:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try57() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg30:=x, Macro:="Targets.Tail")
failure = Err.Number
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Private Function Try58() As String
Dim x As Variant
Dim y As Variant
Dim n As Long
Dim m As Integer
Dim answer As Variant
Dim failure As Long
Dim lateObject As Object
Set lateObject = Application
x = CLng(1)
y = CLng(2)
n = 1
m = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run(Arg01:=x, Macro:="Targets.Whole")
failure = Err.Number
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(x) & ":" & CStr(y) & ":" & CStr(n) & ":" & CStr(m) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "early_reverse=" & Try0() & vbLf
report = report & "early_alias=" & Try1() & vbLf
report = report & "early_forced_second=" & Try2() & vbLf
report = report & "early_forced_first=" & Try3() & vbLf
report = report & "early_hole=" & Try4() & vbLf
report = report & "early_optional_reverse=" & Try5() & vbLf
report = report & "early_long=" & Try6() & vbLf
report = report & "early_variant=" & Try7() & vbLf
report = report & "early_integer=" & Try8() & vbLf
report = report & "early_rvalue=" & Try9() & vbLf
report = report & "early_extra=" & Try10() & vbLf
report = report & "early_missing_and_extra=" & Try11() & vbLf
report = report & "early_tail_hole=" & Try12() & vbLf
report = report & "early_tail30=" & Try13() & vbLf
report = report & "object_reverse=" & Try14() & vbLf
report = report & "object_alias=" & Try15() & vbLf
report = report & "object_forced_second=" & Try16() & vbLf
report = report & "object_forced_first=" & Try17() & vbLf
report = report & "object_hole=" & Try18() & vbLf
report = report & "object_optional_reverse=" & Try19() & vbLf
report = report & "object_long=" & Try20() & vbLf
report = report & "object_variant=" & Try21() & vbLf
report = report & "object_integer=" & Try22() & vbLf
report = report & "object_rvalue=" & Try23() & vbLf
report = report & "object_extra=" & Try24() & vbLf
report = report & "object_missing_and_extra=" & Try25() & vbLf
report = report & "object_tail_hole=" & Try26() & vbLf
report = report & "object_tail30=" & Try27() & vbLf
report = report & "object_unknown=" & Try28() & vbLf
report = report & "with_object_reverse=" & Try29() & vbLf
report = report & "with_object_alias=" & Try30() & vbLf
report = report & "with_object_forced_second=" & Try31() & vbLf
report = report & "with_object_forced_first=" & Try32() & vbLf
report = report & "with_object_hole=" & Try33() & vbLf
report = report & "with_object_optional_reverse=" & Try34() & vbLf
report = report & "with_object_long=" & Try35() & vbLf
report = report & "with_object_variant=" & Try36() & vbLf
report = report & "with_object_integer=" & Try37() & vbLf
report = report & "with_object_rvalue=" & Try38() & vbLf
report = report & "with_object_extra=" & Try39() & vbLf
report = report & "with_object_missing_and_extra=" & Try40() & vbLf
report = report & "with_object_tail_hole=" & Try41() & vbLf
report = report & "with_object_tail30=" & Try42() & vbLf
report = report & "with_object_unknown=" & Try43() & vbLf
report = report & "object_function_reverse=" & Try44() & vbLf
report = report & "object_function_alias=" & Try45() & vbLf
report = report & "object_function_forced_second=" & Try46() & vbLf
report = report & "object_function_forced_first=" & Try47() & vbLf
report = report & "object_function_hole=" & Try48() & vbLf
report = report & "object_function_optional_reverse=" & Try49() & vbLf
report = report & "object_function_long=" & Try50() & vbLf
report = report & "object_function_variant=" & Try51() & vbLf
report = report & "object_function_integer=" & Try52() & vbLf
report = report & "object_function_rvalue=" & Try53() & vbLf
report = report & "object_function_extra=" & Try54() & vbLf
report = report & "object_function_missing_and_extra=" & Try55() & vbLf
report = report & "object_function_tail_hole=" & Try56() & vbLf
report = report & "object_function_tail30=" & Try57() & vbLf
report = report & "object_function_unknown=" & Try58() & vbLf
Probe = report
End Function

