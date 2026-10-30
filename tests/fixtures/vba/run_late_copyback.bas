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
Option Explicit
Public Function Whole(ByRef value As Long) As String
Driver.Trace = Driver.Trace & "T"
value = 99
Whole = "body"
End Function
Public Function Pair(ByRef value As Long, ByRef second As Variant) As String
Driver.Trace = Driver.Trace & "T"
value = 99
second = 88
Pair = "pair"
End Function
Public Function OptionalWhole(ByRef value As Long, Optional ByRef extra As Variant) As String
Driver.Trace = Driver.Trace & "T"
value = 99
OptionalWhole = CStr(IsMissing(extra))
End Function
Public Function ParamWhole(ByRef value As Long, ParamArray rest() As Variant) As String
Driver.Trace = Driver.Trace & "T"
value = 99
ParamWhole = CStr(UBound(rest))
End Function
Public Function ByValWhole(ByVal value As Long) As String
Driver.Trace = Driver.Trace & "T"
value = 99
ByValWhole = "value"
End Function
Public Function AliasValue(ByRef first As Variant, ByRef second As Variant) As String
Driver.Trace = Driver.Trace & "T"
first = 10
second = first + 1
AliasValue = CStr(first) & ":" & CStr(second)
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Function EarlyApp() As Excel.Application
Set EarlyApp = Application
End Function
Public Function LateApp() As Object
Set LateApp = Application
End Function
Private Function Try0() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try1() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try2() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try3() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try4() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try5() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try6() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try7() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try8() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try9() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try10() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try11() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try12() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try13() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try14() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try15() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try16() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try17() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try18() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try19() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try20() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try21() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try22() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try23() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try24() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try25() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try26() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try27() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try28() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try29() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try30() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try31() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try32() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try33() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = earlyObject.Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try34() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try35() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try36() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try37() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try38() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try39() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try40() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try41() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try42() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try43() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try44() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try45() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try46() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try47() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try48() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try49() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try50() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try51() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try52() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try53() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try54() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try55() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try56() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try57() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try58() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try59() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try60() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try61() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try62() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try63() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try63 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try64() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try64 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try65() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try65 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try66() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try66 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try67() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = variantObject.Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try67 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try68() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try68 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try69() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try69 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try70() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try70 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try71() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try71 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try72() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try72 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try73() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try73 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try74() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try74 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try75() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try75 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try76() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try76 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try77() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try77 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try78() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try78 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try79() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try79 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try80() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try80 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try81() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try81 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try82() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try82 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try83() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try83 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try84() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = EarlyApp().Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try84 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try85() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try85 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try86() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try86 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try87() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held)
failure = Err.Number
Err.Clear
Try87 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try88() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try88 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try89() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try89 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try90() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try90 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try91() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try91 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try92() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try92 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try93() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Whole")
failure = Err.Number
Err.Clear
Try93 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try94() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try94 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try95() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.Pair", held)
failure = Err.Number
Err.Clear
Try95 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try96() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try96 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try97() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try97 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try98() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try98 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try99() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try99 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try100() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try100 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try101() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = LateApp().Run("Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try101 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try102() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try102 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try103() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try103 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try104() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try104 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try105() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try105 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try106() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try106 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try107() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try107 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try108() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", (held), 2)
End With
failure = Err.Number
Err.Clear
Try108 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try109() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole", CLng(1), 2)
End With
failure = Err.Number
Err.Clear
Try109 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try110() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try110 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try111() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Pair", held)
End With
failure = Err.Number
Err.Clear
Try111 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try112() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.Pair", held)
End With
failure = Err.Number
Err.Clear
Try112 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try113() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.OptionalWhole", held)
End With
failure = Err.Number
Err.Clear
Try113 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try114() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.OptionalWhole", held)
End With
failure = Err.Number
Err.Clear
Try114 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try115() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.ParamWhole", held, 2, 3)
End With
failure = Err.Number
Err.Clear
Try115 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try116() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.ParamWhole", held, 2, 3)
End With
failure = Err.Number
Err.Clear
Try116 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try117() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.ByValWhole", held)
End With
failure = Err.Number
Err.Clear
Try117 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try118() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run("Targets.AliasValue", held, held)
End With
failure = Err.Number
Err.Clear
Try118 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try119() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try119 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try120() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try120 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try121() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held)
End With
failure = Err.Number
Err.Clear
Try121 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try122() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try122 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try123() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try123 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try124() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", held, 2)
End With
failure = Err.Number
Err.Clear
Try124 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try125() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", (held), 2)
End With
failure = Err.Number
Err.Clear
Try125 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try126() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole", CLng(1), 2)
End With
failure = Err.Number
Err.Clear
Try126 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try127() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Whole")
End With
failure = Err.Number
Err.Clear
Try127 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try128() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Pair", held)
End With
failure = Err.Number
Err.Clear
Try128 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try129() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Pair", held)
End With
failure = Err.Number
Err.Clear
Try129 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try130() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.OptionalWhole", held)
End With
failure = Err.Number
Err.Clear
Try130 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try131() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.OptionalWhole", held)
End With
failure = Err.Number
Err.Clear
Try131 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try132() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.ParamWhole", held, 2, 3)
End With
failure = Err.Number
Err.Clear
Try132 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try133() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.ParamWhole", held, 2, 3)
End With
failure = Err.Number
Err.Clear
Try133 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try134() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.ByValWhole", held)
End With
failure = Err.Number
Err.Clear
Try134 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try135() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.AliasValue", held, held)
End With
failure = Err.Number
Err.Clear
Try135 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try136() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held)
failure = Err.Number
Err.Clear
Try136 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try137() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held)
failure = Err.Number
Err.Clear
Try137 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try138() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held)
failure = Err.Number
Err.Clear
Try138 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try139() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try139 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try140() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try140 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try141() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Integer
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", held, 2)
failure = Err.Number
Err.Clear
Try141 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try142() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", (held), 2)
failure = Err.Number
Err.Clear
Try142 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try143() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole", CLng(1), 2)
failure = Err.Number
Err.Clear
Try143 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try144() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Whole")
failure = Err.Number
Err.Clear
Try144 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try145() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Pair", held)
failure = Err.Number
Err.Clear
Try145 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try146() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Pair", held)
failure = Err.Number
Err.Clear
Try146 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try147() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try147 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try148() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.OptionalWhole", held)
failure = Err.Number
Err.Clear
Try148 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try149() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try149 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try150() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Long
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ParamWhole", held, 2, 3)
failure = Err.Number
Err.Clear
Try150 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try151() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.ByValWhole", held)
failure = Err.Number
Err.Clear
Try151 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Private Function Try152() As String
Dim answer As Variant
Dim failure As Long
Dim earlyObject As Excel.Application
Dim lateObject As Object
Dim variantObject As Variant
Set earlyObject = Application
Set lateObject = Application
Set variantObject = Application
Dim held As Variant
held = CLng(1)
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.AliasValue", held, held)
failure = Err.Number
Err.Clear
Try152 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "direct_long=" & Try0() & vbLf
report = report & "direct_variant=" & Try1() & vbLf
report = report & "direct_integer=" & Try2() & vbLf
report = report & "direct_long_extra=" & Try3() & vbLf
report = report & "direct_variant_extra=" & Try4() & vbLf
report = report & "direct_integer_extra=" & Try5() & vbLf
report = report & "direct_forced_extra=" & Try6() & vbLf
report = report & "direct_literal_extra=" & Try7() & vbLf
report = report & "direct_missing=" & Try8() & vbLf
report = report & "direct_pair_missing_variant=" & Try9() & vbLf
report = report & "direct_pair_missing_long=" & Try10() & vbLf
report = report & "direct_optional_variant=" & Try11() & vbLf
report = report & "direct_optional_long=" & Try12() & vbLf
report = report & "direct_param_variant=" & Try13() & vbLf
report = report & "direct_param_long=" & Try14() & vbLf
report = report & "direct_byval_variant=" & Try15() & vbLf
report = report & "direct_alias=" & Try16() & vbLf
report = report & "early_variable_long=" & Try17() & vbLf
report = report & "early_variable_variant=" & Try18() & vbLf
report = report & "early_variable_integer=" & Try19() & vbLf
report = report & "early_variable_long_extra=" & Try20() & vbLf
report = report & "early_variable_variant_extra=" & Try21() & vbLf
report = report & "early_variable_integer_extra=" & Try22() & vbLf
report = report & "early_variable_forced_extra=" & Try23() & vbLf
report = report & "early_variable_literal_extra=" & Try24() & vbLf
report = report & "early_variable_missing=" & Try25() & vbLf
report = report & "early_variable_pair_missing_variant=" & Try26() & vbLf
report = report & "early_variable_pair_missing_long=" & Try27() & vbLf
report = report & "early_variable_optional_variant=" & Try28() & vbLf
report = report & "early_variable_optional_long=" & Try29() & vbLf
report = report & "early_variable_param_variant=" & Try30() & vbLf
report = report & "early_variable_param_long=" & Try31() & vbLf
report = report & "early_variable_byval_variant=" & Try32() & vbLf
report = report & "early_variable_alias=" & Try33() & vbLf
report = report & "object_variable_long=" & Try34() & vbLf
report = report & "object_variable_variant=" & Try35() & vbLf
report = report & "object_variable_integer=" & Try36() & vbLf
report = report & "object_variable_long_extra=" & Try37() & vbLf
report = report & "object_variable_variant_extra=" & Try38() & vbLf
report = report & "object_variable_integer_extra=" & Try39() & vbLf
report = report & "object_variable_forced_extra=" & Try40() & vbLf
report = report & "object_variable_literal_extra=" & Try41() & vbLf
report = report & "object_variable_missing=" & Try42() & vbLf
report = report & "object_variable_pair_missing_variant=" & Try43() & vbLf
report = report & "object_variable_pair_missing_long=" & Try44() & vbLf
report = report & "object_variable_optional_variant=" & Try45() & vbLf
report = report & "object_variable_optional_long=" & Try46() & vbLf
report = report & "object_variable_param_variant=" & Try47() & vbLf
report = report & "object_variable_param_long=" & Try48() & vbLf
report = report & "object_variable_byval_variant=" & Try49() & vbLf
report = report & "object_variable_alias=" & Try50() & vbLf
report = report & "variant_variable_long=" & Try51() & vbLf
report = report & "variant_variable_variant=" & Try52() & vbLf
report = report & "variant_variable_integer=" & Try53() & vbLf
report = report & "variant_variable_long_extra=" & Try54() & vbLf
report = report & "variant_variable_variant_extra=" & Try55() & vbLf
report = report & "variant_variable_integer_extra=" & Try56() & vbLf
report = report & "variant_variable_forced_extra=" & Try57() & vbLf
report = report & "variant_variable_literal_extra=" & Try58() & vbLf
report = report & "variant_variable_missing=" & Try59() & vbLf
report = report & "variant_variable_pair_missing_variant=" & Try60() & vbLf
report = report & "variant_variable_pair_missing_long=" & Try61() & vbLf
report = report & "variant_variable_optional_variant=" & Try62() & vbLf
report = report & "variant_variable_optional_long=" & Try63() & vbLf
report = report & "variant_variable_param_variant=" & Try64() & vbLf
report = report & "variant_variable_param_long=" & Try65() & vbLf
report = report & "variant_variable_byval_variant=" & Try66() & vbLf
report = report & "variant_variable_alias=" & Try67() & vbLf
report = report & "early_function_long=" & Try68() & vbLf
report = report & "early_function_variant=" & Try69() & vbLf
report = report & "early_function_integer=" & Try70() & vbLf
report = report & "early_function_long_extra=" & Try71() & vbLf
report = report & "early_function_variant_extra=" & Try72() & vbLf
report = report & "early_function_integer_extra=" & Try73() & vbLf
report = report & "early_function_forced_extra=" & Try74() & vbLf
report = report & "early_function_literal_extra=" & Try75() & vbLf
report = report & "early_function_missing=" & Try76() & vbLf
report = report & "early_function_pair_missing_variant=" & Try77() & vbLf
report = report & "early_function_pair_missing_long=" & Try78() & vbLf
report = report & "early_function_optional_variant=" & Try79() & vbLf
report = report & "early_function_optional_long=" & Try80() & vbLf
report = report & "early_function_param_variant=" & Try81() & vbLf
report = report & "early_function_param_long=" & Try82() & vbLf
report = report & "early_function_byval_variant=" & Try83() & vbLf
report = report & "early_function_alias=" & Try84() & vbLf
report = report & "object_function_long=" & Try85() & vbLf
report = report & "object_function_variant=" & Try86() & vbLf
report = report & "object_function_integer=" & Try87() & vbLf
report = report & "object_function_long_extra=" & Try88() & vbLf
report = report & "object_function_variant_extra=" & Try89() & vbLf
report = report & "object_function_integer_extra=" & Try90() & vbLf
report = report & "object_function_forced_extra=" & Try91() & vbLf
report = report & "object_function_literal_extra=" & Try92() & vbLf
report = report & "object_function_missing=" & Try93() & vbLf
report = report & "object_function_pair_missing_variant=" & Try94() & vbLf
report = report & "object_function_pair_missing_long=" & Try95() & vbLf
report = report & "object_function_optional_variant=" & Try96() & vbLf
report = report & "object_function_optional_long=" & Try97() & vbLf
report = report & "object_function_param_variant=" & Try98() & vbLf
report = report & "object_function_param_long=" & Try99() & vbLf
report = report & "object_function_byval_variant=" & Try100() & vbLf
report = report & "object_function_alias=" & Try101() & vbLf
report = report & "with_object_long=" & Try102() & vbLf
report = report & "with_object_variant=" & Try103() & vbLf
report = report & "with_object_integer=" & Try104() & vbLf
report = report & "with_object_long_extra=" & Try105() & vbLf
report = report & "with_object_variant_extra=" & Try106() & vbLf
report = report & "with_object_integer_extra=" & Try107() & vbLf
report = report & "with_object_forced_extra=" & Try108() & vbLf
report = report & "with_object_literal_extra=" & Try109() & vbLf
report = report & "with_object_missing=" & Try110() & vbLf
report = report & "with_object_pair_missing_variant=" & Try111() & vbLf
report = report & "with_object_pair_missing_long=" & Try112() & vbLf
report = report & "with_object_optional_variant=" & Try113() & vbLf
report = report & "with_object_optional_long=" & Try114() & vbLf
report = report & "with_object_param_variant=" & Try115() & vbLf
report = report & "with_object_param_long=" & Try116() & vbLf
report = report & "with_object_byval_variant=" & Try117() & vbLf
report = report & "with_object_alias=" & Try118() & vbLf
report = report & "with_application_long=" & Try119() & vbLf
report = report & "with_application_variant=" & Try120() & vbLf
report = report & "with_application_integer=" & Try121() & vbLf
report = report & "with_application_long_extra=" & Try122() & vbLf
report = report & "with_application_variant_extra=" & Try123() & vbLf
report = report & "with_application_integer_extra=" & Try124() & vbLf
report = report & "with_application_forced_extra=" & Try125() & vbLf
report = report & "with_application_literal_extra=" & Try126() & vbLf
report = report & "with_application_missing=" & Try127() & vbLf
report = report & "with_application_pair_missing_variant=" & Try128() & vbLf
report = report & "with_application_pair_missing_long=" & Try129() & vbLf
report = report & "with_application_optional_variant=" & Try130() & vbLf
report = report & "with_application_optional_long=" & Try131() & vbLf
report = report & "with_application_param_variant=" & Try132() & vbLf
report = report & "with_application_param_long=" & Try133() & vbLf
report = report & "with_application_byval_variant=" & Try134() & vbLf
report = report & "with_application_alias=" & Try135() & vbLf
report = report & "callbyname_long=" & Try136() & vbLf
report = report & "callbyname_variant=" & Try137() & vbLf
report = report & "callbyname_integer=" & Try138() & vbLf
report = report & "callbyname_long_extra=" & Try139() & vbLf
report = report & "callbyname_variant_extra=" & Try140() & vbLf
report = report & "callbyname_integer_extra=" & Try141() & vbLf
report = report & "callbyname_forced_extra=" & Try142() & vbLf
report = report & "callbyname_literal_extra=" & Try143() & vbLf
report = report & "callbyname_missing=" & Try144() & vbLf
report = report & "callbyname_pair_missing_variant=" & Try145() & vbLf
report = report & "callbyname_pair_missing_long=" & Try146() & vbLf
report = report & "callbyname_optional_variant=" & Try147() & vbLf
report = report & "callbyname_optional_long=" & Try148() & vbLf
report = report & "callbyname_param_variant=" & Try149() & vbLf
report = report & "callbyname_param_long=" & Try150() & vbLf
report = report & "callbyname_byval_variant=" & Try151() & vbLf
report = report & "callbyname_alias=" & Try152() & vbLf
Probe = report
End Function

