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
Public Function Reader(ByRef first As Variant) As String
Driver.Trace = Driver.Trace & "T" & CStr(Err.Number)
Reader = CStr(first) & ":" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function
Public Function Pair(ByRef first As Variant, ByRef second As Variant) As String
Driver.Trace = Driver.Trace & "T" & CStr(Err.Number)
first = first + 10
second = first + 1
Pair = CStr(first) & ":" & CStr(second) & ":" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function
Public Function Whole(ByRef first As Long) As String
Driver.Trace = Driver.Trace & "T" & CStr(Err.Number)
Whole = CStr(first) & ":" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
first = 99
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Which As String
Public MoveGlobal As Boolean
Public ChangeError As Boolean
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
Driver.Trace = Driver.Trace & "D"
If MoveGlobal Then
Driver.Held = 5
Driver.TypedHeld = 5
End If
If ChangeError Then
On Error Resume Next
Err.Raise 55, "Default", "default"
End If
Value = Which
End Property


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Held As Variant
Public TypedHeld As Long
Private Function Try0() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try1() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try2() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try3() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try4() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try5() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try6() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try7() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try8() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try9() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try10() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try11() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try12() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try13() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try14() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try15() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try16() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try17() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try18() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try19() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try20() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try21() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try22() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try23() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try24() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try25() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try26() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try27() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try28() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try29() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try30() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try31() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try32() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try33() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try34() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try35() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try36() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try37() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try38() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try39() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try40() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try41() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try42() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try43() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try44() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try45() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try46() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try47() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try48() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try49() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try50() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try51() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try52() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try53() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try54() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try55() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try56() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try57() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try58() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try59() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try60() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try61() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try62() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try63() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = lateObject.Run(chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try63 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try64() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try64 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try65() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try65 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try66() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try66 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try67() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try67 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try68() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try68 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try69() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try69 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try70() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try70 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try71() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try71 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try72() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try72 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try73() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try73 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try74() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try74 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try75() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try75 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try76() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try76 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try77() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try77 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try78() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try78 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try79() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try79 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try80() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try80 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try81() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try81 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try82() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try82 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try83() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try83 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try84() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try84 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try85() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try85 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try86() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try86 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try87() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try87 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try88() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try88 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try89() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try89 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try90() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try90 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try91() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try91 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try92() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try92 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try93() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try93 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try94() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, Held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try94 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try95() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With lateObject
answer = .Run(chosen, TypedHeld)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try95 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try96() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try96 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try97() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try97 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try98() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try98 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try99() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try99 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try100() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try100 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try101() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try101 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try102() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try102 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try103() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try103 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try104() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try104 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try105() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try105 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try106() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try106 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try107() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try107 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try108() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try108 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try109() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try109 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try110() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try110 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try111() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try111 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try112() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try112 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try113() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try113 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try114() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try114 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try115() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try115 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try116() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try116 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try117() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try117 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try118() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try118 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try119() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = False
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try119 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try120() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try120 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try121() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try121 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try122() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try122 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try123() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = False
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try123 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try124() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Reader"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try124 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try125() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Pair"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try125 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try126() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, Held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try126 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try127() As String
Dim chosen As MacroName
Dim lateObject As Object
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Set chosen = New MacroName
Set lateObject = Application
chosen.Which = "Targets.Whole"
chosen.MoveGlobal = True
chosen.ChangeError = True
Held = CLng(1)
TypedHeld = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, chosen, TypedHeld)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try127 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(Held) & ":" & CStr(TypedHeld) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "early_prior0_move0_err0_read=" & Try0() & vbLf
report = report & "early_prior0_move0_err0_pair=" & Try1() & vbLf
report = report & "early_prior0_move0_err0_variant_long=" & Try2() & vbLf
report = report & "early_prior0_move0_err0_typed_long=" & Try3() & vbLf
report = report & "early_prior0_move0_err1_read=" & Try4() & vbLf
report = report & "early_prior0_move0_err1_pair=" & Try5() & vbLf
report = report & "early_prior0_move0_err1_variant_long=" & Try6() & vbLf
report = report & "early_prior0_move0_err1_typed_long=" & Try7() & vbLf
report = report & "early_prior0_move1_err0_read=" & Try8() & vbLf
report = report & "early_prior0_move1_err0_pair=" & Try9() & vbLf
report = report & "early_prior0_move1_err0_variant_long=" & Try10() & vbLf
report = report & "early_prior0_move1_err0_typed_long=" & Try11() & vbLf
report = report & "early_prior0_move1_err1_read=" & Try12() & vbLf
report = report & "early_prior0_move1_err1_pair=" & Try13() & vbLf
report = report & "early_prior0_move1_err1_variant_long=" & Try14() & vbLf
report = report & "early_prior0_move1_err1_typed_long=" & Try15() & vbLf
report = report & "early_prior1_move0_err0_read=" & Try16() & vbLf
report = report & "early_prior1_move0_err0_pair=" & Try17() & vbLf
report = report & "early_prior1_move0_err0_variant_long=" & Try18() & vbLf
report = report & "early_prior1_move0_err0_typed_long=" & Try19() & vbLf
report = report & "early_prior1_move0_err1_read=" & Try20() & vbLf
report = report & "early_prior1_move0_err1_pair=" & Try21() & vbLf
report = report & "early_prior1_move0_err1_variant_long=" & Try22() & vbLf
report = report & "early_prior1_move0_err1_typed_long=" & Try23() & vbLf
report = report & "early_prior1_move1_err0_read=" & Try24() & vbLf
report = report & "early_prior1_move1_err0_pair=" & Try25() & vbLf
report = report & "early_prior1_move1_err0_variant_long=" & Try26() & vbLf
report = report & "early_prior1_move1_err0_typed_long=" & Try27() & vbLf
report = report & "early_prior1_move1_err1_read=" & Try28() & vbLf
report = report & "early_prior1_move1_err1_pair=" & Try29() & vbLf
report = report & "early_prior1_move1_err1_variant_long=" & Try30() & vbLf
report = report & "early_prior1_move1_err1_typed_long=" & Try31() & vbLf
report = report & "object_prior0_move0_err0_read=" & Try32() & vbLf
report = report & "object_prior0_move0_err0_pair=" & Try33() & vbLf
report = report & "object_prior0_move0_err0_variant_long=" & Try34() & vbLf
report = report & "object_prior0_move0_err0_typed_long=" & Try35() & vbLf
report = report & "object_prior0_move0_err1_read=" & Try36() & vbLf
report = report & "object_prior0_move0_err1_pair=" & Try37() & vbLf
report = report & "object_prior0_move0_err1_variant_long=" & Try38() & vbLf
report = report & "object_prior0_move0_err1_typed_long=" & Try39() & vbLf
report = report & "object_prior0_move1_err0_read=" & Try40() & vbLf
report = report & "object_prior0_move1_err0_pair=" & Try41() & vbLf
report = report & "object_prior0_move1_err0_variant_long=" & Try42() & vbLf
report = report & "object_prior0_move1_err0_typed_long=" & Try43() & vbLf
report = report & "object_prior0_move1_err1_read=" & Try44() & vbLf
report = report & "object_prior0_move1_err1_pair=" & Try45() & vbLf
report = report & "object_prior0_move1_err1_variant_long=" & Try46() & vbLf
report = report & "object_prior0_move1_err1_typed_long=" & Try47() & vbLf
report = report & "object_prior1_move0_err0_read=" & Try48() & vbLf
report = report & "object_prior1_move0_err0_pair=" & Try49() & vbLf
report = report & "object_prior1_move0_err0_variant_long=" & Try50() & vbLf
report = report & "object_prior1_move0_err0_typed_long=" & Try51() & vbLf
report = report & "object_prior1_move0_err1_read=" & Try52() & vbLf
report = report & "object_prior1_move0_err1_pair=" & Try53() & vbLf
report = report & "object_prior1_move0_err1_variant_long=" & Try54() & vbLf
report = report & "object_prior1_move0_err1_typed_long=" & Try55() & vbLf
report = report & "object_prior1_move1_err0_read=" & Try56() & vbLf
report = report & "object_prior1_move1_err0_pair=" & Try57() & vbLf
report = report & "object_prior1_move1_err0_variant_long=" & Try58() & vbLf
report = report & "object_prior1_move1_err0_typed_long=" & Try59() & vbLf
report = report & "object_prior1_move1_err1_read=" & Try60() & vbLf
report = report & "object_prior1_move1_err1_pair=" & Try61() & vbLf
report = report & "object_prior1_move1_err1_variant_long=" & Try62() & vbLf
report = report & "object_prior1_move1_err1_typed_long=" & Try63() & vbLf
report = report & "with_object_prior0_move0_err0_read=" & Try64() & vbLf
report = report & "with_object_prior0_move0_err0_pair=" & Try65() & vbLf
report = report & "with_object_prior0_move0_err0_variant_long=" & Try66() & vbLf
report = report & "with_object_prior0_move0_err0_typed_long=" & Try67() & vbLf
report = report & "with_object_prior0_move0_err1_read=" & Try68() & vbLf
report = report & "with_object_prior0_move0_err1_pair=" & Try69() & vbLf
report = report & "with_object_prior0_move0_err1_variant_long=" & Try70() & vbLf
report = report & "with_object_prior0_move0_err1_typed_long=" & Try71() & vbLf
report = report & "with_object_prior0_move1_err0_read=" & Try72() & vbLf
report = report & "with_object_prior0_move1_err0_pair=" & Try73() & vbLf
report = report & "with_object_prior0_move1_err0_variant_long=" & Try74() & vbLf
report = report & "with_object_prior0_move1_err0_typed_long=" & Try75() & vbLf
report = report & "with_object_prior0_move1_err1_read=" & Try76() & vbLf
report = report & "with_object_prior0_move1_err1_pair=" & Try77() & vbLf
report = report & "with_object_prior0_move1_err1_variant_long=" & Try78() & vbLf
report = report & "with_object_prior0_move1_err1_typed_long=" & Try79() & vbLf
report = report & "with_object_prior1_move0_err0_read=" & Try80() & vbLf
report = report & "with_object_prior1_move0_err0_pair=" & Try81() & vbLf
report = report & "with_object_prior1_move0_err0_variant_long=" & Try82() & vbLf
report = report & "with_object_prior1_move0_err0_typed_long=" & Try83() & vbLf
report = report & "with_object_prior1_move0_err1_read=" & Try84() & vbLf
report = report & "with_object_prior1_move0_err1_pair=" & Try85() & vbLf
report = report & "with_object_prior1_move0_err1_variant_long=" & Try86() & vbLf
report = report & "with_object_prior1_move0_err1_typed_long=" & Try87() & vbLf
report = report & "with_object_prior1_move1_err0_read=" & Try88() & vbLf
report = report & "with_object_prior1_move1_err0_pair=" & Try89() & vbLf
report = report & "with_object_prior1_move1_err0_variant_long=" & Try90() & vbLf
report = report & "with_object_prior1_move1_err0_typed_long=" & Try91() & vbLf
report = report & "with_object_prior1_move1_err1_read=" & Try92() & vbLf
report = report & "with_object_prior1_move1_err1_pair=" & Try93() & vbLf
report = report & "with_object_prior1_move1_err1_variant_long=" & Try94() & vbLf
report = report & "with_object_prior1_move1_err1_typed_long=" & Try95() & vbLf
report = report & "callbyname_prior0_move0_err0_read=" & Try96() & vbLf
report = report & "callbyname_prior0_move0_err0_pair=" & Try97() & vbLf
report = report & "callbyname_prior0_move0_err0_variant_long=" & Try98() & vbLf
report = report & "callbyname_prior0_move0_err0_typed_long=" & Try99() & vbLf
report = report & "callbyname_prior0_move0_err1_read=" & Try100() & vbLf
report = report & "callbyname_prior0_move0_err1_pair=" & Try101() & vbLf
report = report & "callbyname_prior0_move0_err1_variant_long=" & Try102() & vbLf
report = report & "callbyname_prior0_move0_err1_typed_long=" & Try103() & vbLf
report = report & "callbyname_prior0_move1_err0_read=" & Try104() & vbLf
report = report & "callbyname_prior0_move1_err0_pair=" & Try105() & vbLf
report = report & "callbyname_prior0_move1_err0_variant_long=" & Try106() & vbLf
report = report & "callbyname_prior0_move1_err0_typed_long=" & Try107() & vbLf
report = report & "callbyname_prior0_move1_err1_read=" & Try108() & vbLf
report = report & "callbyname_prior0_move1_err1_pair=" & Try109() & vbLf
report = report & "callbyname_prior0_move1_err1_variant_long=" & Try110() & vbLf
report = report & "callbyname_prior0_move1_err1_typed_long=" & Try111() & vbLf
report = report & "callbyname_prior1_move0_err0_read=" & Try112() & vbLf
report = report & "callbyname_prior1_move0_err0_pair=" & Try113() & vbLf
report = report & "callbyname_prior1_move0_err0_variant_long=" & Try114() & vbLf
report = report & "callbyname_prior1_move0_err0_typed_long=" & Try115() & vbLf
report = report & "callbyname_prior1_move0_err1_read=" & Try116() & vbLf
report = report & "callbyname_prior1_move0_err1_pair=" & Try117() & vbLf
report = report & "callbyname_prior1_move0_err1_variant_long=" & Try118() & vbLf
report = report & "callbyname_prior1_move0_err1_typed_long=" & Try119() & vbLf
report = report & "callbyname_prior1_move1_err0_read=" & Try120() & vbLf
report = report & "callbyname_prior1_move1_err0_pair=" & Try121() & vbLf
report = report & "callbyname_prior1_move1_err0_variant_long=" & Try122() & vbLf
report = report & "callbyname_prior1_move1_err0_typed_long=" & Try123() & vbLf
report = report & "callbyname_prior1_move1_err1_read=" & Try124() & vbLf
report = report & "callbyname_prior1_move1_err1_pair=" & Try125() & vbLf
report = report & "callbyname_prior1_move1_err1_variant_long=" & Try126() & vbLf
report = report & "callbyname_prior1_move1_err1_typed_long=" & Try127() & vbLf
Probe = report
End Function

