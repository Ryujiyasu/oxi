VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Function Clean(ByRef held As Variant) As String
held = 99
Driver.Trace = Driver.Trace & "C" & CStr(Err.Number)
Clean = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Function Clean(ByRef held As Variant) As String
held = 99
Driver.Trace = Driver.Trace & "C" & CStr(Err.Number)
Clean = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function


Attribute VB_Name = "Targets"
Public Function Clean(ByRef held As Variant) As String
held = 99
Driver.Trace = Driver.Trace & "C" & CStr(Err.Number)
Clean = CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get App() As Object
On Error Resume Next
Driver.Trace = Driver.Trace & "G"
Err.Raise 55, "Getter", "getter"
Set App = Application
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Run(ByVal macroName As Variant, ByRef held As Variant) As String
held = 99
Driver.Trace = Driver.Trace & "S" & CStr(Err.Number)
Run = "shadow:" & CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Public Function Arg() As Variant
On Error Resume Next
Trace = Trace & "A"
Err.Raise 66, "Argument", "argument"
Arg = 1
End Function
Public Function NameArg() As String
On Error Resume Next
Trace = Trace & "N"
Err.Raise 77, "Name", "name"
NameArg = "Targets.Clean"
End Function
Private Function Try0() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try0 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try1() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try1 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try2() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try2 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try3() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try3 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try4() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try4 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try5() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try5 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try6() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try6 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try7() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try7 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try8() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try8 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try9() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Targets.Clean", Arg())
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try9 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try10() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run(NameArg(), held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try10 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try11() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run(NameArg(), Arg())
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try11 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try12() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run("Targets.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try12 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try13() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run("Targets.Clean", Arg())
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try13 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try14() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run(NameArg(), held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try14 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try15() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run(NameArg(), Arg())
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try15 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try16() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try16 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try17() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try17 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try18() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try18 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try19() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try19 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try20() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try20 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try21() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try21 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try22() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try22 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try23() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try23 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try24() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try24 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try25() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try25 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try26() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try26 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try27() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try27 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try28() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try28 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try29() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try29 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try30() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try30 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try31() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try31 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try32() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = factoryObj.App.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try32 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try33() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = factoryObj.App.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try33 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try34() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = factoryObj.App.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try34 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try35() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = factoryObj.App.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try35 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try36() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = factoryObj.App.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try36 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try37() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = factoryObj.App.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try37 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try38() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = factoryObj.App.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try38 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try39() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = factoryObj.App.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try39 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try40() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = shadowObj.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try40 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try41() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = shadowObj.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try41 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try42() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = shadowObj.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try42 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try43() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = shadowObj.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try43 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try44() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = shadowObj.Run("Targets.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try44 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try45() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = shadowObj.Run("Targets.Clean", Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try45 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try46() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = shadowObj.Run(NameArg(), held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try46 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try47() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = shadowObj.Run(NameArg(), Arg())
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try47 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try48() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try48 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try49() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try49 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try50() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("ThisWorkbook.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try50 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try51() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run("ThisWorkbook.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try51 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try52() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run("ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try52 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try53() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run("ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try53 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try54() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try54 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try55() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "ThisWorkbook.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try55 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try56() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Application.Run("Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try56 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try57() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Application.Run("Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try57 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try58() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
With Application
answer = .Run("Sheet1.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try58 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try59() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
With Application
answer = .Run("Sheet1.Clean", held)
End With
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try59 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try60() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = Run("Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try60 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try61() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = Run("Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try61 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try62() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Clear
answer = CallByName(Application, "Run", VbMethod, "Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try62 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Private Function Try63() As String
Dim held As Variant
Dim answer As Variant
Dim failure As Long
Dim failureSource As String
Dim failureDescription As String
Dim factoryObj As Factory
Dim shadowObj As Shadow
Set factoryObj = New Factory
Set shadowObj = New Shadow
held = 1
Trace = ""
On Error Resume Next
Err.Raise 42, "Outer", "before"
answer = CallByName(Application, "Run", VbMethod, "Sheet1.Clean", held)
failure = Err.Number
failureSource = Err.Source
failureDescription = Err.Description
Err.Clear
Try63 = CStr(failure) & "|" & CStr(answer) & "|" & CStr(held) & "|" & failureSource & "|" & failureDescription & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "run_clear_ref=" & Try0() & vbLf
report = report & "run_clear_arg=" & Try1() & vbLf
report = report & "run_clear_name=" & Try2() & vbLf
report = report & "run_clear_both=" & Try3() & vbLf
report = report & "run_prior_ref=" & Try4() & vbLf
report = report & "run_prior_arg=" & Try5() & vbLf
report = report & "run_prior_name=" & Try6() & vbLf
report = report & "run_prior_both=" & Try7() & vbLf
report = report & "with_clear_ref=" & Try8() & vbLf
report = report & "with_clear_arg=" & Try9() & vbLf
report = report & "with_clear_name=" & Try10() & vbLf
report = report & "with_clear_both=" & Try11() & vbLf
report = report & "with_prior_ref=" & Try12() & vbLf
report = report & "with_prior_arg=" & Try13() & vbLf
report = report & "with_prior_name=" & Try14() & vbLf
report = report & "with_prior_both=" & Try15() & vbLf
report = report & "global_clear_ref=" & Try16() & vbLf
report = report & "global_clear_arg=" & Try17() & vbLf
report = report & "global_clear_name=" & Try18() & vbLf
report = report & "global_clear_both=" & Try19() & vbLf
report = report & "global_prior_ref=" & Try20() & vbLf
report = report & "global_prior_arg=" & Try21() & vbLf
report = report & "global_prior_name=" & Try22() & vbLf
report = report & "global_prior_both=" & Try23() & vbLf
report = report & "callbyname_clear_ref=" & Try24() & vbLf
report = report & "callbyname_clear_arg=" & Try25() & vbLf
report = report & "callbyname_clear_name=" & Try26() & vbLf
report = report & "callbyname_clear_both=" & Try27() & vbLf
report = report & "callbyname_prior_ref=" & Try28() & vbLf
report = report & "callbyname_prior_arg=" & Try29() & vbLf
report = report & "callbyname_prior_name=" & Try30() & vbLf
report = report & "callbyname_prior_both=" & Try31() & vbLf
report = report & "getter_clear_ref=" & Try32() & vbLf
report = report & "getter_clear_arg=" & Try33() & vbLf
report = report & "getter_clear_name=" & Try34() & vbLf
report = report & "getter_clear_both=" & Try35() & vbLf
report = report & "getter_prior_ref=" & Try36() & vbLf
report = report & "getter_prior_arg=" & Try37() & vbLf
report = report & "getter_prior_name=" & Try38() & vbLf
report = report & "getter_prior_both=" & Try39() & vbLf
report = report & "shadow_clear_ref=" & Try40() & vbLf
report = report & "shadow_clear_arg=" & Try41() & vbLf
report = report & "shadow_clear_name=" & Try42() & vbLf
report = report & "shadow_clear_both=" & Try43() & vbLf
report = report & "shadow_prior_ref=" & Try44() & vbLf
report = report & "shadow_prior_arg=" & Try45() & vbLf
report = report & "shadow_prior_name=" & Try46() & vbLf
report = report & "shadow_prior_both=" & Try47() & vbLf
report = report & "run_ThisWorkbook_clear=" & Try48() & vbLf
report = report & "run_ThisWorkbook_prior=" & Try49() & vbLf
report = report & "with_ThisWorkbook_clear=" & Try50() & vbLf
report = report & "with_ThisWorkbook_prior=" & Try51() & vbLf
report = report & "global_ThisWorkbook_clear=" & Try52() & vbLf
report = report & "global_ThisWorkbook_prior=" & Try53() & vbLf
report = report & "callbyname_ThisWorkbook_clear=" & Try54() & vbLf
report = report & "callbyname_ThisWorkbook_prior=" & Try55() & vbLf
report = report & "run_Sheet1_clear=" & Try56() & vbLf
report = report & "run_Sheet1_prior=" & Try57() & vbLf
report = report & "with_Sheet1_clear=" & Try58() & vbLf
report = report & "with_Sheet1_prior=" & Try59() & vbLf
report = report & "global_Sheet1_clear=" & Try60() & vbLf
report = report & "global_Sheet1_prior=" & Try61() & vbLf
report = report & "callbyname_Sheet1_clear=" & Try62() & vbLf
report = report & "callbyname_Sheet1_prior=" & Try63() & vbLf
Probe = report
End Function

