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
Public Function Echo(ParamArray args() As Variant) As String
Driver.Trace = Driver.Trace & "B"
Echo = CStr(UBound(args) + 1)
End Function


Attribute VB_Name = "Driver"
Option Explicit
Public Trace As String
Private Function GetName() As String
Trace = Trace & "N"
GetName = "Targets.Echo"
End Function
Private Function GetArgument() As Variant
Trace = Trace & "A"
GetArgument = CLng(7)
End Function
Private Function Try0() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Macro:="Targets.Echo", Arg2:=9)
failure = Err.Number
Try0 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try1() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Macro:="Targets.Echo", Arg30:=9)
failure = Err.Number
Try1 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try2() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run("Targets.Echo", Arg2:=9)
failure = Err.Number
Try2 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try3() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Arg1:=7, Macro:="Targets.Echo")
failure = Err.Number
Try3 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try4() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Unknown:="Targets.Echo", Arg1:=7)
failure = Err.Number
Try4 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try5() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Macro:="Targets.Echo", Arg31:=7)
failure = Err.Number
Try5 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try6() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Arg1:=7)
failure = Err.Number
Try6 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try7() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Empty, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
failure = Err.Number
Try7 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try8() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(CVErr(2042), 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
failure = Err.Number
Try8 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try9() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Nothing, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
failure = Err.Number
Try9 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try10() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Arg1:=GetArgument(), Macro:=GetName())
failure = Err.Number
Try10 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try11() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Application.Run(Arg1:=GetArgument(), Macro:=GetName())
failure = Err.Number
Try11 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try12() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
value = Run(Arg1:=GetArgument(), Macro:=GetName())
failure = Err.Number
Try12 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try13() As String
Dim value As Variant
Dim failure As Long
Trace = ""
On Error Resume Next
Err.Clear
With Application
value = .Run(Arg1:=GetArgument(), Macro:=GetName())
End With
failure = Err.Number
Try13 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try14() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(GetName(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument(), GetArgument())
failure = Err.Number
Try14 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try15() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Macro:="Targets.Echo", Arg0:=7)
failure = Err.Number
Try15 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Private Function Try16() As String
Dim value As Variant
Dim failure As Long
Dim app As Object
Set app = Application
Trace = ""
On Error Resume Next
Err.Clear
value = app.Run(Macro:="Targets.Echo", Arg01:=7)
failure = Err.Number
Try16 = CStr(failure) & "|" & TypeName(value) & "|" & CStr(value) & "|" & Trace
End Function
Public Function Probe() As String
Dim report As String
report = report & "arg2=" & Try0() & vbLf
report = report & "arg30=" & Try1() & vbLf
report = report & "mixed=" & Try2() & vbLf
report = report & "fold=" & Try3() & vbLf
report = report & "unknown=" & Try4() & vbLf
report = report & "arg31_name=" & Try5() & vbLf
report = report & "missing_macro=" & Try6() & vbLf
report = report & "empty31=" & Try7() & vbLf
report = report & "error31=" & Try8() & vbLf
report = report & "nothing31=" & Try9() & vbLf
report = report & "late_eval=" & Try10() & vbLf
report = report & "early_eval=" & Try11() & vbLf
report = report & "global_eval=" & Try12() & vbLf
report = report & "with_eval=" & Try13() & vbLf
report = report & "extra31_eval=" & Try14() & vbLf
report = report & "arg0_name=" & Try15() & vbLf
report = report & "arg01_name=" & Try16() & vbLf
Probe = report
End Function

